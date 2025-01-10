package services;

import static services.ExcelUtility.*;
import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Objects;

import classes.ImportBaseDto;
import interfaces.ExcelImportConfig;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ConvertExcel {

	/**
	 * Reads an Excel file and converts its content into a list of DTOs (Data Transfer Objects).
	 * The method processes the file, extracts the data based on the provided configuration, 
	 * and returns a list of objects of the specified type.
	 * 
	 * @param <T> The type of the DTO to which the data will be mapped. It must extend {@link ImportBaseDto}.
	 * @param file The byte array representing the Excel file to be read.
	 * @param excelImportConfig A list of {@link ExcelImportConfig} objects, which define how the Excel data should be mapped.
	 * @param readingDto The class type of the DTO to which data will be mapped.
	 * 
	 * @return A list of DTOs populated with data extracted from the Excel file.
	 * @throws RuntimeException If an error occurs during the Excel file reading or processing.
	 */
	public static <T extends ImportBaseDto> List<T> readExcel(byte[] file, List<ExcelImportConfig> excelImportConfig, Class<T> readingDto) {
		try (InputStream is = new ByteArrayInputStream(file);
			 XSSFWorkbook workbook = new XSSFWorkbook(is)) {
			checkConfigList(excelImportConfig);
			
			Iterator<Sheet> sheetIterator = workbook.sheetIterator();
			//TODO Possibile gestione di sheet multipli
			Sheet sheet = sheetIterator.next();

			Iterator<Row> rowIterator = sheet.rowIterator();
			return startRead(excelImportConfig, readingDto, rowIterator,null);
		} catch (Exception e) {
			throw new RuntimeException(e.getMessage(),e);
		} finally {
			//TODO LOGGER
		}
	}

	/**
	 * Reads the rows of an Excel sheet and converts them into a list of DTOs (Data Transfer Objects).
	 * This method processes the rows based on the provided Excel import configuration, mapping the data 
	 * from the Excel cells to the appropriate DTO class.
	 * 
	 * @param <T> The type of the DTO to which the data will be mapped. It must extend {@link ImportBaseDto}.
	 * @param excelImportConfig A list of {@link ExcelImportConfig} objects that define how the Excel data should be mapped to the DTO.
	 * @param readingDto The class type of the DTO to which data will be mapped.
	 * @param rowIterator An iterator for the rows of the Excel sheet.
	 * @param first The first row to read, if any (can be null, in which case the first row from the iterator is used).
	 * 
	 * @return A list of DTOs populated with data extracted from the Excel sheet rows.
	 * @throws RuntimeException If any required field is missing or an error occurs during data reading.
	 */
	private static <T extends ImportBaseDto> List<T> startRead(List<ExcelImportConfig> excelImportConfig,Class<T> readingDto, Iterator<Row> rowIterator, Row first) {
		int startColumn = excelImportConfig.parallelStream().filter(config -> config.getStartColumn() != null).findFirst().map(ExcelImportConfig::getStartColumn).orElse(1);
		Row firstRow = first != null ? first : rowIterator.next();
		
		Iterator<Cell> titleIterator = firstRow.cellIterator();
		Cell forceCell = skipCell(startColumn, titleIterator);  // (prima cella utile)
		Map<Integer, ExcelImportConfig> orderColumnExcel = checkTitlesExcel(titleIterator,excelImportConfig, forceCell);
	
		List<T> response = readRowExcel(rowIterator, orderColumnExcel,startColumn,readingDto);
		
		checkRequiredFields(excelImportConfig, response);
		
		return response;
	}

	/**
	 * Reads the rows of an Excel sheet and maps each row to a DTO (Data Transfer Object) based on the provided configuration.
	 * This method processes the rows starting from a specified row, extracts the data from the columns, 
	 * and converts each row into an object of the specified DTO type.
	 * 
	 * @param <T> The type of the DTO to which the data will be mapped. It must extend {@link ImportBaseDto}.
	 * @param rowIterator An iterator for the rows of the Excel sheet.
	 * @param orderColumnExcel A map that associates Excel column indices with the corresponding {@link ExcelImportConfig}.
	 * @param startColumn The starting column index from which to begin reading the data in each row.
	 * @param readingDto The class type of the DTO to which the data will be mapped.
	 * 
	 * @return A list of DTOs populated with data extracted from the Excel rows.
	 * @throws RuntimeException If an error occurs during the reading or mapping of rows, including issues with instantiating the DTO.
	 */
	private static <T extends ImportBaseDto> List<T> readRowExcel(Iterator<Row> rowIterator, Map<Integer, ExcelImportConfig> orderColumnExcel,int startColumn,Class<T> readingDto) {
		List<T> listaImport = new ArrayList<>();
		int firstRow = orderColumnExcel.values().parallelStream().filter(config -> config.getStartRow() != null).findFirst().map(ExcelImportConfig::getStartRow).orElse(1);
		while(rowIterator.hasNext()) {
			Row row = rowIterator.next();
			
			if(row.getRowNum() < firstRow) continue;
			
			T importDto = null;
			try {
				importDto = readingDto.getDeclaredConstructor().newInstance();
				importDto.setNrRow(row.getRowNum());
				if (readRow(row, orderColumnExcel,startColumn, importDto)) continue;
				listaImport.add(importDto);
			} catch (InstantiationException | IllegalAccessException | IllegalArgumentException
					| InvocationTargetException | NoSuchMethodException | SecurityException e) {
				throw new RuntimeException("Error in import data");
			}
		}
		return listaImport;
	}
	
	/**
	 * Processes a single row from the Excel sheet, reads the cell data, and maps it to the provided DTO (Data Transfer Object).
	 * This method iterates over the cells of the row starting from a specified column and uses the provided configuration 
	 * to map the data to the DTO.
	 * 
	 * @param <T> The type of the DTO to which the data will be mapped.
	 * @param row The row from the Excel sheet to be processed.
	 * @param orderColumnExcel A map that associates Excel column indices with the corresponding {@link ExcelImportConfig}.
	 * @param startColumn The index of the first column from which to begin reading the data in the row.
	 * @param readingDto The DTO to which the data from the row will be mapped.
	 * 
	 * @return A boolean indicating whether the row was processed successfully (true) or if it was skipped (false).
	 */
	private static <T> boolean readRow(Row row,Map<Integer, ExcelImportConfig> orderColumnExcel,int startColumn,T readingDto) {
		int maxIndexCell = orderColumnExcel.keySet().stream().max(Integer::compare).orElse((int)row.getLastCellNum());
		Iterator<Cell> cellIterator = row.cellIterator();
		Cell skipCell = skipCell(startColumn, cellIterator);
		return readCellsExcel(orderColumnExcel, maxIndexCell, readingDto, cellIterator, skipCell);
	}
	
	/**
	 * Processes the cells of a single row in the Excel sheet and maps the data to the provided DTO (Data Transfer Object).
	 * This method iterates over the cells, checks the corresponding configuration for each column, 
	 * and updates the DTO fields based on the cell values. 
	 * The method also handles skipping cells, verifying data types, and checking for special row statuses.
	 * 
	 * @param <T> The type of the DTO to which the cell data will be mapped. 
	 * @param orderColumnExcel A map that associates the Excel column index with the corresponding {@link ExcelImportConfig} for each column.
	 * @param maxIndexCell The maximum index of the cell to process in the row. This ensures we process cells up to this index.
	 * @param readingDto The DTO object that will be populated with the data from the cells.
	 * @param cellIterator An iterator for the cells in the row.
	 * @param skipCell The cell to skip if needed (based on the configuration).
	 * 
	 * @return A boolean indicating whether the row should be skipped. It returns true if the row is marked for skipping, 
	 *         otherwise returns false.
	 */
	private static <T> boolean readCellsExcel(Map<Integer, ExcelImportConfig> orderColumnExcel, int maxIndexCell, T readingDto, Iterator<Cell> cellIterator, Cell skipCell) {
		List<Field> declaredFields = Arrays.asList(readingDto.getClass().getDeclaredFields());
		List<ExcelRowStatus> statusCells = new ArrayList<>();
		ExcelRowStatus status = ExcelRowStatus.EMPTY_ROW;
		boolean isFirstIteration = true;
		while(cellIterator.hasNext() && maxIndexCell >= 0) {
			Cell cell = null;
			if(isFirstIteration  && Objects.nonNull(skipCell)) {
				cell = skipCell;
				isFirstIteration = false;
			} else {
				cell = cellIterator.next();
			}
			if(ExcelRowStatus.SPECIAL.equals(status)) break;
			try {
				ExcelImportConfig label = orderColumnExcel.get(cell.getColumnIndex());
				if(label == null || label.getId() == null) {
					continue;
				}
				if(!cell.getCellType().equals(label.getColumnType())) {
					//TODO LOGGER
				}
				
			    status = processCell(readingDto, getField(declaredFields,label.getColumnOrder(),label.getColumnTargetField()), cell, label);
			    statusCells.add(status);
			} finally {
				maxIndexCell--;
			}
		};
		return ExcelRowStatus.SPECIAL.equals(status) || statusCells.stream().allMatch(s -> ExcelRowStatus.EMPTY_ROW.equals(s));
	}
	
	/**
	 * Processes a single cell in the Excel sheet, evaluates its value based on the cell type (string, numeric, boolean, or formula), 
	 * and maps the value to the corresponding field of the provided DTO (Data Transfer Object).
	 * This method handles the different types of Excel cells and applies the necessary logic for each type.
	 * 
	 * @param <T> The type of the DTO that the cell data will be mapped to.
	 * @param readingDto The DTO object that will be populated with the cell data.
	 * @param field The field of the DTO to which the cell value will be mapped.
	 * @param cell The cell in the Excel sheet that contains the value to be processed.
	 * @param label The configuration for the column, including whether the cell value is required and how it should be handled.
	 * 
	 * @return An {@link ExcelRowStatus} indicating the result of processing the cell. This status could represent whether the row is empty, 
	 *         contains errors, or has been successfully processed.
	 */
	private static <T> ExcelRowStatus processCell(T readingDto, Field field, Cell cell, ExcelImportConfig label) {
        return switch (cell.getCellType()) {
            case STRING ->
                    checkValueAndProcess(field, readingDto, handleCell(cell, label.getRequired(), (c) -> c.getStringCellValue().trim()));
            case NUMERIC ->
                    checkValueAndProcess(field, readingDto, handleCell(cell, label.getRequired(), (c) -> DateUtil.isCellDateFormatted(cell) ? c.getDateCellValue() : c.getNumericCellValue()));
            case BOOLEAN ->
                    checkValueAndProcess(field, readingDto, handleCell(cell, label.getRequired(), (c) -> c.getBooleanCellValue()));
            case FORMULA -> checkValueAndProcess(field, readingDto, handleCell(cell, label.getRequired(), (c) -> {
                FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
                CellValue evaluate = evaluator.evaluate(cell);
                CellType cellType = evaluate.getCellType();
                return CellType.NUMERIC.equals(cellType) ? c.getNumericCellValue() : c.getStringCellValue();
            }));
            default -> ExcelRowStatus.EMPTY_ROW;
        };
	}
	
	/**
	 * Verifies and processes the value of a cell, setting the value in the corresponding field of the entity if the value is valid.
	 * If the value is special, the method returns a special row status, otherwise it maps the value to the field in the entity.
	 *
	 * @param <T> The type of the cell value to be processed.
	 * @param <E> The type of the entity (e.g., a DTO) in which the value will be mapped.
	 * @param field The field of the entity that will be updated with the cell value.
	 * @param entity The entity that contains the field to be updated.
	 * @param value The value of the cell to be processed and mapped to the entity's field.
	 * 
	 * @return An {@link ExcelRowStatus} value that represents the status of the row after processing the cell.
	 *         It can return a special status if the value is special, or an empty row status if the value is null.
	 */
	private static <T,E> ExcelRowStatus checkValueAndProcess(Field field, E entity,T value) {
		if(field == null)
			return ExcelRowStatus.EMPTY_ROW;
		Field fieldSpecial = getFieldSpecial(Arrays.asList(entity.getClass().getDeclaredFields()),value.toString());
		if(fieldSpecial != null)
			return ExcelRowStatus.SPECIAL;
		ExcelRowStatus status = ExcelRowStatus.EMPTY_ROW;
		if(value != null)
			status = setField(field, entity, value);
		
		return status;
	}
	
	

//=========================================================================================================
	/**
	 * METODI PER LA LETTURA DI UN FILE EXCEL CON PARAMETRI DI RICERCA
	 */
//=========================================================================================================
	
	/**
	 * Reads an Excel file and extracts data into a list of DTOs while also retrieving search parameters
	 * defined by a separate configuration. This method handles both importing data and extracting parameters
	 * from the provided Excel sheet.
	 *
	 * @param <T> The type of DTO that will be used to map the rows of the Excel file (must extend {@link ImportBaseDto}).
	 * @param file A byte array representing the Excel file to be processed.
	 * @param excelImportConfig A list of configurations defining how to map the rows of the Excel file to DTOs.
	 * @param parametriRicercaConf A list of configurations defining how to extract search parameters from the Excel file.
	 * @param readingDto The class of the DTO type to map the rows of the Excel file.
	 * 
	 * @return A {@link Entry} where the key is containing the extracted
	 *         search parameters and the value is a list of DTOs containing the data from the Excel file.
	 * 
	 * @throws RuntimeException if any error occurs during the reading and processing of the Excel file.
	 */
	public static <T extends ImportBaseDto> Entry<? extends ImportBaseDto,List<T>> readExcelWithGeneralsTable(byte[] file, List<ExcelImportConfig> excelImportConfig,List<ExcelImportConfig> parametriRicercaConf, Class<T> readingDto) {
		checkConfigList(excelImportConfig);
		checkConfigList(parametriRicercaConf);
		int startRow = excelImportConfig.parallelStream().filter(config -> config.getStartRow() != null).findFirst().map(ExcelImportConfig::getStartRow).orElse(1);
		
		try (InputStream is = new ByteArrayInputStream(file);
			 XSSFWorkbook workbook = new XSSFWorkbook(is)) {
			Iterator<Sheet> sheetIterator = workbook.sheetIterator();
			//TODO Possibile gestione di sheet multipli
			Sheet sheet = sheetIterator.next();

			Iterator<Row> rowIterator = sheet.rowIterator();
			ImportBaseDto firstTable = new ImportBaseDto();

			Row firstRow = startReadParametriRicerca(parametriRicercaConf,startRow, rowIterator,firstTable);
			
			List<T> response = startRead(excelImportConfig, readingDto, rowIterator, firstRow);
			return Map.entry(firstTable, response);
		} catch (Exception e) {
			throw new RuntimeException(e.getMessage(),e);
		}
	}

	/**
	 * Reads a row from the Excel file and processes the cells according to the given configuration for the search parameters.
	 * The search parameters are extracted and stored in the provided {@link ImportBaseDto} object.
	 * The method processes rows starting from a given row number and stops once it reaches the start row for data processing.
	 *
	 * @param generalsTableConf A list of configurations defining how to map the search parameters from the Excel rows.
	 * @param startRow The row number from which the Excel data reading should begin (inclusive).
	 * @param rowIterator An iterator over the rows of the Excel sheet.
	 * @param generalsTable The DTO where the extracted search parameters will be stored.
	 * 
	 * @return The first row containing data for the main processing, or null if no row is found.
	 * 
	 * @throws RuntimeException if an error occurs during reading or processing the Excel file.
	 */
	private static Row startReadParametriRicerca(List<ExcelImportConfig> generalsTableConf, int startRow, Iterator<Row> rowIterator, ImportBaseDto generalsTable) {
		List<Field> fields = Arrays.asList(generalsTable.getClass().getDeclaredFields());
		
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			if(row.getRowNum() >= startRow-1) return row;
			Iterator<Cell> cellIterator = row.cellIterator();
			handleGeneralsTableCells(generalsTableConf, generalsTable, fields, cellIterator);
		}
		return null;
	}

	/**
	 * Processes the cells in the current row of the Excel sheet to extract and set the search parameters.
	 * The method iterates through the cells, identifies the title (header) cells, and then maps the values to
	 * the corresponding fields in the {@link ImportBaseDto} object.
	 * The title row helps in determining which columns represent search parameters.
	 *
	 * @param generalsTableConf A list of configurations that define how to map the search parameter columns from the Excel sheet.
	 * @param generalsTable The DTO where the extracted search parameters will be stored.
	 * @param fields A list of fields in the {@link ImportBaseDto} that will be populated with values from the cells.
	 * @param cellIterator An iterator over the cells in the row being processed.
	 *
	 * @throws RuntimeException if an error occurs while processing a cell or mapping a field.
	 */
	private static void handleGeneralsTableCells(List<ExcelImportConfig> generalsTableConf, ImportBaseDto generalsTable, List<Field> fields, Iterator<Cell> cellIterator) {
		boolean title = true;
		Field field = null;
		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			if(title) {
				final String value = CellType.STRING.equals(cell.getCellType()) ? cell.getStringCellValue() : null;
				ExcelImportConfig label = getLabelByColumnName(generalsTableConf, value);
				if(label == null) {
				continue;
				}
				generalsTableConf.removeIf(l -> l.getId().equals(label.getId()));
				field = getField(fields, label.getColumnOrder(), label.getColumnTargetField());
				title = false;
				continue;
			}

			if(field == null) continue;

			if(processGeneralsTableCell(generalsTable, field, cell))
				break;
		}
	}

	/**
	 * Processes an individual cell from the Excel sheet and sets its value to the appropriate field in the {@link ImportBaseDto}.
	 * The method determines the type of the cell (e.g., STRING, NUMERIC) and applies the correct logic to retrieve the value. 
	 * It then sets the field in the DTO if the value is valid.
	 *
	 * @param generalsTable The {@link ImportBaseDto} object where the cell value will be set.
	 * @param field The field in the {@link ImportBaseDto} that corresponds to the value in the cell.
	 * @param cell The Excel cell containing the value to be processed and assigned to the DTO field.
	 * 
	 * @return {@code true} if the cell was processed and the field was set; {@code false} otherwise.
	 *         The method returns {@code true} if the field was successfully updated, and {@code false} if no valid data was found.
	 */
	private static boolean processGeneralsTableCell(ImportBaseDto generalsTable, Field field, Cell cell) {

        return switch (cell.getCellType()) {
            case STRING ->
                    ExcelRowStatus.skip.contains(setField(field, generalsTable, handleCell(cell, false, Cell::getStringCellValue)));
            case NUMERIC ->
                    ExcelRowStatus.skip.contains(setField(field, generalsTable, handleCell(cell, false, (c) -> DateUtil.isCellDateFormatted(cell) ? c.getDateCellValue() : c.getNumericCellValue())));
            default -> false;
        };
    }
}
