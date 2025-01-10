package services;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import classes.ComplexExport;
import classes.ExportBaseTable;
import classes.ReportExport;
import classes.ReportImportDto;
import enums.StatusRowEnum;
import interfaces.ExcelImportConfig;
import interfaces.ExportBaseInterface;
import interfaces.ExportSimple;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static services.ExcelUtility.*;

public abstract class ExportExcel {
	
	private final static  String DEFAULT_OK_MESSAGE = "IMPORT OK";
	
	/**
	 * This method generates a report in an existing Excel file by processing the provided data and configurations.
	 * It loads the provided Excel file (in byte array format), processes the data based on the provided configuration 
	 * and report details, then writes the updated data into the Excel file and returns the byte array of the modified file.
	 *
	 * @param file The Excel file in byte array format that will be updated with the report.
	 * @param configExcel A list of {@link ExcelImportConfig} objects which define the configuration 
	 *                    for importing data (such as start row and column positions).
	 * @param reports A list of {@link ReportImportDto} objects containing the report data to be processed and applied to the Excel file.
	 * @return A byte array representing the modified Excel file with the generated report.
	 * 
	 * @throws RuntimeException If an error occurs while processing the Excel file or generating the report, 
	 *                          an exception will be thrown with the error message.
	 */
	public static byte[] generateReportOnExistingFile(byte[] file, List<ExcelImportConfig> configExcel, List<ReportImportDto> reports) {
		try(InputStream is = new ByteArrayInputStream(file);
			 XSSFWorkbook workbook = new XSSFWorkbook(is)){
			
			Iterator<Sheet> sheetIterator = workbook.sheetIterator();
			//TODO Possibile gestione di sheet multipli
			Sheet sheet = sheetIterator.next();
			
			int startRow = configExcel.parallelStream().filter(config -> config.getStartRow() != null).findFirst().map(ExcelImportConfig::getStartRow).orElse(1);
			int startColumn = configExcel.parallelStream().filter(config -> config.getStartColumn() != null).findFirst().map(ExcelImportConfig::getStartColumn).orElse(1);
			
			Iterator<Row> rowIterator = sheet.rowIterator();
			Row forceRow = skipRow(startRow, rowIterator);
            assert forceRow != null;
            Iterator<Cell> cellIterator = forceRow.cellIterator();
			Cell skipCell = skipCell(startColumn, cellIterator);
			
			Map<Integer, ExcelImportConfig> orderConfig = checkTitlesExcel(cellIterator, configExcel, skipCell);
			
			Set<Integer> keySet = orderConfig.keySet();
			Integer lastCellIndex = keySet.parallelStream().max(Integer::compareTo).orElse(null);
			
			while(rowIterator.hasNext()){
				Row row = rowIterator.next();
				if(Objects.isNull(lastCellIndex)) {
					lastCellIndex = (int)row.getLastCellNum();
				}
				removeOldCommentAndCellStyle(row);
				
				if(reports != null && !reports.isEmpty()) {
					List<ReportImportDto> errorRow = reports.parallelStream().filter(err -> err.getRowIndex() == row.getRowNum()).collect(Collectors.toList());
					if(!errorRow.isEmpty()) {
						handleReport(workbook, sheet, orderConfig, row, errorRow,lastCellIndex);
						continue;
					}
				}
				if(emptyRow(lastCellIndex, row.cellIterator()))break;
				createReportCell(DEFAULT_OK_MESSAGE, workbook, row, lastCellIndex,IndexedColors.LIGHT_GREEN);
			};
			sheet.autoSizeColumn(lastCellIndex+1);
			
			ByteArrayOutputStream response = new ByteArrayOutputStream();
			workbook.write(response);
			return response.toByteArray();
		} catch (Exception e) {
			throw new RuntimeException(e.getMessage(),e);
		}	
	}

	/**
	 * This method removes any existing comments and resets the cell style for each cell in the given row.
	 * It iterates over all cells in the row, removes any attached comments, and sets the cell's background color 
	 * to white by resetting the fill foreground color of the cell style.
	 *
	 * @param row The row from which to remove comments and reset cell styles. The method will process each cell in the row.
	 */
	private static void removeOldCommentAndCellStyle(Row row) {
		row.forEach(c -> {
			c.removeCellComment();
			c.getCellStyle().setFillForegroundColor(IndexedColors.WHITE.getIndex());
		});
	}

	/**
	 * This method creates a report cell in the given row of the Excel sheet. It inserts the provided message 
	 * into the cell and applies a custom style to the cell. The cell is placed in the column immediately 
	 * following the last cell in the row.
	 * 
	 * If the target cell does not already exist, it is created. The method also sets the value of the cell 
	 * to the provided message (`message`) and applies a background color using the provided `color`.
	 * 
	 * @param message The message to be inserted into the newly created cell.
	 * @param workbook The workbook that contains the sheet, used to apply the custom style.
	 * @param row The row where the new report cell will be created.
	 * @param lastCellIndex The index of the last cell in the row, used to determine the position for the new cell.
	 * @param color The background color to be applied to the newly created cell. It uses the `IndexedColors` enumeration.
	 */
	private static void createReportCell(String message, XSSFWorkbook workbook, Row row, int lastCellIndex, IndexedColors color) {
		Cell cell = row.getCell(lastCellIndex+1);
		if(cell == null) {
			cell = row.createCell(lastCellIndex+1,CellType.STRING);
		}
		cell.setCellValue(new XSSFRichTextString(message));
		cell.setCellStyle(customStyle(workbook, null, color));
	}

	/**
	 * This method checks whether a row is empty based on its cells. It iterates over the cells in the row and 
	 * determines if all cells up to the specified `lastCellIndex` are empty (blank). A row is considered empty 
	 * if all cells in the specified range are blank.
	 * 
	 * @param lastCellIndex The index of the last cell in the row to check. The method will iterate over cells 
	 *                      from the first to this index.
	 * @param cellIterator The iterator that traverses the cells in the row.
	 * 
	 * @return {@code true} if all cells up to the `lastCellIndex` are blank, {@code false} otherwise.
	 */
	private static boolean emptyRow(int lastCellIndex, Iterator<Cell> cellIterator) {
		boolean emptyRow = true;
		while(cellIterator.hasNext() && lastCellIndex >= 0) {
			Cell cell = cellIterator.next();
			if(!CellType.BLANK.equals(cell.getCellType()))
				emptyRow = false;
		}
		return emptyRow;
	}

	/**
	 * This method processes a list of error reports and updates the Excel sheet accordingly. For each error report, 
	 * it adds a message to the corresponding row and cell, and applies appropriate styles based on the error's 
	 * state and column index. If the report indicates an error, the method marks the cell with a background color.
	 * 
	 * The following actions are performed for each report:
	 * 1. If the report status is "OK" (enums.StatoRowEnum.OK), a success message is inserted into the cell with a green background.
	 * 2. If the column index in the report is null, a message is inserted into the cell with a yellow background.
	 * 3. If a valid column index is provided, the corresponding cell is marked with an appropriate style.
	 *
	 * @param workbook The workbook containing the sheet to be updated.
	 * @param sheet The sheet within the workbook that will be modified.
	 * @param orderConfig A map containing the configuration for each column in the Excel sheet.
	 * @param row The row where the error reports will be applied.
	 * @param errorRow A list of {@link ReportImportDto} objects containing the error messages and additional details.
	 * @param lastCell The index of the last cell in the row, used to determine the position for the error message.
	 */
	private static void handleReport(XSSFWorkbook workbook, Sheet sheet, Map<Integer, ExcelImportConfig> orderConfig,Row row, List<ReportImportDto> errorRow, int lastCell) {
		errorRow.forEach(report -> {
			if(StatusRowEnum.OK.equals(report.getStatus())) {
				createReportCell(report.getMessage(), workbook, row, lastCell, IndexedColors.LIGHT_GREEN);
			} else if(report.getColumnIndex() == null) {
				createReportCell(report.getMessage(), workbook, row, lastCell, StatusRowEnum.ERROR.equals(report.getStatus()) ? IndexedColors.RED : IndexedColors.LIGHT_YELLOW);
			} else {
				markCell(workbook, sheet, orderConfig, report, row , lastCell);
			}
		});
	}

	/**
	 * This method marks a specific cell in the row with a warning style based on the given error report and column configuration.
	 * It iterates over the column configuration map (`orderConfig`) to find the column index that matches the error's column index.
	 * Once the correct column is found, it applies a warning style to the corresponding cell in the row.
	 * 
	 * @param workbook The workbook containing the sheet to be updated.
	 * @param sheet The sheet within the workbook where the cell will be marked.
	 * @param orderConfig A map of column configurations, where the key is the column index and the value is the column's configuration.
	 * @param error The error report (`ReportImportDto`) containing the column index and other error details to be marked.
	 * @param row The row where the cell will be updated.
	 */
	private static void markCell(XSSFWorkbook workbook, Sheet sheet, Map<Integer, ExcelImportConfig> orderConfig, ReportImportDto error, Row row, int lastCell) {
		for(Entry<Integer, ExcelImportConfig> it : orderConfig.entrySet()) {
			ExcelImportConfig value = it.getValue();
			Integer key = it.getKey();
			if(value.getColumnOrder().equals(error.getColumnIndex())) {
				setWarningCell(workbook, sheet, error, row, row.getCell(key));
				createReportCell(error.getMessage(), workbook, row, lastCell, StatusRowEnum.ERROR.equals(error.getStatus()) ? IndexedColors.RED : IndexedColors.LIGHT_YELLOW);
				break;
			}
		}
	}

	/**
	 * This method applies a warning style to a specific cell and adds a comment to it based on the provided error report.
	 * It sets the cell's style to a yellow background (using the `IndexedColors.LIGHT_YELLOW` color) and adds a comment 
	 * with the error message to the cell. The comment helps to indicate why the cell was marked, providing additional context 
	 * for the error or warning.
	 * 
	 * @param workbook The workbook containing the sheet to be updated.
	 * @param sheet The sheet within the workbook where the cell will be updated.
	 * @param error The error report (`ReportImportDto`) containing the message to be added as a comment to the cell.
	 * @param row The row containing the cell to be updated.
	 * @param cell The cell that will be marked with the warning style and the comment.
	 */
	private static void setWarningCell(XSSFWorkbook workbook, Sheet sheet, ReportImportDto error, Row row,	Cell cell) {
		CellStyle warningStyle = customStyle(workbook, cell.getCellStyle(),IndexedColors.LIGHT_YELLOW);
		cell.setCellStyle(warningStyle);
		
		Comment cellComment = getCellComment(workbook, sheet, row, cell);
		cellComment.setString(workbook.getCreationHelper().createRichTextString(error.getMessage()));
	}

	/**
	 * This method creates and returns a custom cell style with a specified background color. 
	 * It applies the provided color to the cell's background and ensures the fill pattern is set to a solid foreground.
	 * 
	 * The style is based on an existing cell style, which is retrieved using the `getCellStyle` method. 
	 * The new style is then modified with the provided color and returned.
	 * 
	 * @param workbook The workbook containing the sheet and the styles.
	 * @param cell The cell whose style will be modified.
	 * @param color The background color to be applied to the cell. This color is from the `IndexedColors` enum.
	 * 
	 * @return The updated `CellStyle` with the specified background color.
	 */
	private static CellStyle customStyle(XSSFWorkbook workbook, CellStyle cell,IndexedColors color) {
		CellStyle cellStyle = getCellStyle(workbook, cell);
		cellStyle.setFillForegroundColor(color.getIndex());
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		return cellStyle;
	}

	/**
	 * This method creates a new `CellStyle` based on an existing one. If a valid `CellStyle` is provided, 
	 * it clones the style from the existing `CellStyle` and returns the cloned style. If no style is provided 
	 * (i.e., the `cellStyle` parameter is `null`), it creates and returns a new, default `CellStyle`.
	 * 
	 * @param workbook The workbook containing the styles.
	 * @param cellStyle The existing `CellStyle` to be cloned. If null, a new default style will be created.
	 * 
	 * @return A new `CellStyle` object, either cloned from the provided style or a default one if no style is given.
	 */
	private static CellStyle getCellStyle(XSSFWorkbook workbook, CellStyle cellStyle) {
		CellStyle resp = workbook.createCellStyle();
		if(cellStyle != null) {
			resp.cloneStyleFrom(cellStyle);
		}
		return resp;
	}

	/**
	 * This method retrieves the comment associated with a specific cell. If the cell does not already have a comment, 
	 * it creates a new comment with a default anchor and returns it.
	 * The comment is associated with the given row and cell in the provided sheet.
	 * 
	 * The created comment is positioned using a `ClientAnchor` to set its location on the sheet relative to the specified cell.
	 * If the cell already has a comment, that existing comment is returned.
	 * 
	 * @param workbook The workbook that contains the sheet and allows creating comments.
	 * @param sheet The sheet that contains the cell.
	 * @param row The row containing the cell, used to position the comment.
	 * @param cell The cell for which the comment is retrieved or created.
	 * 
	 * @return The `Comment` associated with the cell. If no comment exists, a new comment is created and returned.
	 */
	private static Comment getCellComment(XSSFWorkbook workbook, Sheet sheet, Row row, Cell cell) {
		Comment cellComment = cell.getCellComment();
		
		if(cellComment == null) {
			Drawing<?> drawing = sheet.createDrawingPatriarch();
			
			ClientAnchor anchor = workbook.getCreationHelper().createClientAnchor();
			anchor.setCol1(cell.getColumnIndex());
			anchor.setRow1(row.getRowNum());
			anchor.setCol2(cell.getColumnIndex() + 5);
			anchor.setRow2(row.getRowNum() + 5);
			
			cellComment = drawing.createCellComment(anchor);
		}
		return cellComment;
	}
	
//==================================================================================================================//
	


	/**
	 * This class is used for building and generating Excel files with customizable settings. 
	 * It allows for method chaining to set various parameters for the Excel file, including:
	 * 
	 * <ul>
	 *     <li><b>Font name:</b> Set the font name for the entire Excel document.</li>
	 *     <li><b>Background color for table headers:</b> Set the background color for the header cells in tables.</li>
	 *     <li><b>Table orientation:</b> Choose whether the tables should be oriented horizontally or vertically.</li>
	 *     <li><b>Distance between tables:</b> Set the spacing between tables in the Excel sheet.</li>
	 * </ul>
	 * 
	 * The builder pattern allows for clean, flexible, and readable construction of the Excel file with the desired configurations.
	 * 
	 * Example usage:
	 * <pre>
	 * ExportBuilder builder = new ExportBuilder();
	 * builder.setFontName("Arial")
	 *        .setBackgroundColorTitle(IndexedColors.GREY_25_PERCENT)
	 *        .setTableOrientation(Orientation.HORIZONTAL)
	 *        .setDistanceBetweenTables(10);
	 * </pre>
	 */
	public static class ExportBuilder {

//================================================================================================================================
		/**
		 * COMMONS VARIABLES
		 */
//================================================================================================================================
		/**
		 * Configuration settings for the Excel export builder font.
		 * Default is 'Arial' font style
		 */
		private String FONT_NAME = "Arial";

		/**
		 * The background color for the header cells of tables in the Excel sheet.
		 * Default is a light blue color represented by an XSSFColor.
		 */
		private Color backgroundColorHeader = new XSSFColor(new byte[] { (byte) 221, (byte) 235, (byte) 247 }, new DefaultIndexedColorMap());

		/**
		 * The distance between tables in the Excel sheet.
		 * Default is 1, representing the number of rows or space between consecutive tables.
		 */
		private int distanceTable = 1;
		
		/**
		 * The column index where the freeze pane starts.
		 * <p>
		 * If this value is 0, no columns are frozen. 
		 * A value greater than 0 indicates the number of columns to be frozen (fixed on the left).
		 * </p>
		 */
		private int colFreezePane = 0;

		/**
		 * The row index where the freeze pane starts.
		 * <p>
		 * If this value is 0, no rows are frozen. 
		 * A value greater than 0 indicates the number of rows to be frozen (fixed at the top).
		 * </p>
		 */
		private int rowFreezePane = 0;

		
		/**
		 * The builder for the Excel content, responsible for generating the data and formatting.
		 * This will contain logic to build the Excel file itself.
		 */
		private BuilderExcel builderExcel = null;

		/**
		 * The orientation of the tables in the Excel sheet.
		 * Determines whether the tables are arranged vertically or horizontally.
		 * Default value is {@link AlignExcel#VERTICAL}.
		 */
		private AlignExcel orientation = AlignExcel.VERTICAL;

		/**
		 * Default height for header text in the Excel file.
		 * Used for setting the row height for header cells.
		 */
		static final short HEADER_TEXT_HEIGHT = 11;

		/**
		 * Default height for title text in the Excel file.
		 * Used for setting the row height for title cells.
		 */
		static final short TITLE_TEXT_HEIGHT = 12;

		/**
		 * Default height for common text in the Excel file.
		 * Used for setting the row height for general text cells.
		 */
		static final short COMMON_TEXT_HEIGHT = 10;

		/**
		 * Row height for header rows in the Excel sheet.
		 * This value is calculated based on the {@link #HEADER_TEXT_HEIGHT}.
		 */
		static final short ROW_HEADER_HEIGHT = (short) (HEADER_TEXT_HEIGHT * 4);

		/**
		 * Row height for title rows in the Excel sheet.
		 * This value is calculated based on the {@link #TITLE_TEXT_HEIGHT}.
		 */
		static final short ROW_TITLE_HEIGHT = (short) (TITLE_TEXT_HEIGHT * 4);

		/**
		 * Placeholder string for cells that require a formula.
		 */
		private static final String PLACEHOLDER_CELL = "\\?";

		private ExportBuilder() {}
		
		/**
		 * Gets an instance of the {@link ExportBuilder} for constructing an Excel export.
		 * This method provides a way to create a new builder instance.
		 *
		 * @return A new instance of {@link ExportBuilder}.
		 */
		public static ExportBuilder getInstance() {
		    return new ExportBuilder();
		}

		/**
		 * The column index where the freeze pane starts.
		 * <p>
		 * If this value is 0, no columns are frozen. 
		 * A value greater than 0 indicates the number of columns to be frozen (fixed on the left).
		 * </p>
		 * 
		 * This works only with {@link AlignExcel#VERTICAL}.
		 */
		public ExportBuilder setColFreezePane(int col) {
			this.colFreezePane = col;
			return this;
		}
		
		/**
		 * The row index where the freeze pane starts.
		 * <p>
		 * If this value is 0, no rows are frozen. 
		 * A value greater than 0 indicates the number of rows to be frozen (fixed at the top).
		 * </p>
		 * 
		 * This works only with {@link AlignExcel#VERTICAL} and only if you have to print <b>ONE</b> table.
		 */
		public ExportBuilder setRowFreezePane(int row) {
			this.rowFreezePane = row;
			return this;
		}
		
		/**
		 * Sets the font name for the entire Excel document.
		 * This method allows you to specify the font name used in the document.
		 *
		 * @param fontName The font name to set, such as "Arial", "Calibri", etc.
		 * @return The current instance of {@link ExportBuilder} to allow method chaining.
		 */
		public ExportBuilder setFontName(String fontName) {
		    this.FONT_NAME = fontName;
		    return this;
		}

		/**
		 * Sets the background color for the header cells in the Excel sheet.
		 * This method allows you to set the header cell background color using a {@link Color}.
		 *
		 * @param color The color to set as the background for the header cells.
		 * @return The current instance of {@link ExportBuilder} to allow method chaining.
		 */
		public ExportBuilder setBackgroundColorHeader(Color color) {
		    backgroundColorHeader = color;
		    return this;
		}

		/**
		 * Sets the background color for the header cells in the Excel sheet.
		 * This method allows you to set the header cell background color using a byte array 
		 * representing RGB values. The color is mapped using {@link XSSFColor}.
		 *
		 * @param color The RGB values (as a byte array) to set as the background for the header cells.
		 *               Example: { (byte) 221, (byte) 235, (byte) 247 } for a light blue.
		 * @return The current instance of {@link ExportBuilder} to allow method chaining.
		 */
		public ExportBuilder setBackgroundColorHeader(byte[] color) {
		    backgroundColorHeader = new XSSFColor(color, new DefaultIndexedColorMap());
		    return this;
		}

		/**
		 * Sets the distance between tables in the Excel sheet.
		 * This method allows you to define how many rows of space should be placed between consecutive tables.
		 *
		 * @param distance The distance (in number of rows) between tables.
		 * @return The current instance of {@link ExportBuilder} to allow method chaining.
		 */
		public ExportBuilder setDistanceTable(int distance) {
		    this.distanceTable = distance;
		    return this;
		}

		/**
		 * Sets the orientation of the tables in the Excel sheet.
		 * You can choose whether the tables are arranged vertically or horizontally.
		 * 
		 * @param orientation The orientation for the tables, either {@link AlignExcel#VERTICAL} or {@link AlignExcel#HORIZONTAL}.
		 *                    If {@code null} is provided, the current orientation remains unchanged.
		 * @return The current instance of {@link ExportBuilder} to allow method chaining.
		 */
		public ExportBuilder setOrientation(AlignExcel orientation) {
		    if (orientation != null)
		        this.orientation = orientation;
		    return this;
		}

		
//================================================================================================================================
		/**
		 * START GENERATE METHODS
		 */
//================================================================================================================================

		/**
		 * Keeps track of the columns that have already been populated.
		 * This variable is only used when the table orientation is {@link AlignExcel#HORIZONTAL}.
		 */
		private int cellBefore = 0;

		/**
		 * Temporary variable that tracks the columns that were populated previously.
		 * This is used as a copy of the {@link #cellBefore} variable for certain operations.
		 */
		private int cellBeforeCopy = cellBefore;

		/**
		 * List of entities to write to the Excel sheet.
		 * This is used for the final auto-sizing of columns based on the content of these entities.
		 */
		private List<List<? extends ExportSimple>> entities = new ArrayList<>();

		
		/**
		 * Generates a complex Excel report with multiple tables and additional features such as pivot tables or general data.
		 * This method orchestrates the creation of an Excel report by processing the provided data and generating 
		 * corresponding tables, handling special fields, and applying formatting and auto-sizing.
		 * 
		 * @param report The complete representation of the report to be generated. This object contains the necessary 
		 *               data for generating the report, including general data and the tables to be included.
		 * 
		 * @return A byte array representing the generated Excel file. This file contains the formatted data and tables, 
		 *         ready for download or further processing.
		 */
		public byte[] generateReportExcel(ReportExport<? extends ExportBaseTable> report) {
			builderExcel = generateBuilderExcel();
			Optional<? extends ExportBaseTable> datiGenerali = Optional.ofNullable(report.getGeneralities());
			List<ComplexExport<? extends ExportSimple>> dati = report.getData();
			
			BuilderExcel.SheetBuilder sheetBuilder = null;
			List<Integer> specialRows = new ArrayList<>();
			List<SpecialField> specialFieldsGeneralData = null;
			if(datiGenerali.isPresent()) {
				ExportBaseTable obj = datiGenerali.get();
				specialFieldsGeneralData = obj.getSpecialFields();
				sheetBuilder = exGeneralitiesTable(specialFieldsGeneralData, obj,specialRows);
				sheetBuilder.createEmptyRow(distanceTable);
			}
			
			for(ComplexExport<? extends ExportSimple> table : dati) {
				sheetBuilder = exComplexExportForReport(sheetBuilder, specialRows, specialFieldsGeneralData,table);
			};
			if(builderExcel != null)
				builderExcel.autoSizeColumns(getAutoSize());
			
			if(AlignExcel.VERTICAL.equals(orientation)) {
                assert sheetBuilder != null;
                sheetBuilder.freezePane(colFreezePane, rowFreezePane);
            }

            assert builderExcel != null;

            return builderExcel.build();
		}

		/**
		 * Generates the table with general data at the beginning of the Excel sheet.
		 * This method creates a sheet, generates a title, populates the sheet with a vertical key-value table, 
		 * and sets empty rows for special fields if applicable.
		 * 
		 * @param specialFieldsDatiGenerali A list of special fields to handle differently within the sheet.
		 *                                   These fields may require additional processing or formatting.
		 * @param obj The object containing the general data to populate the sheet, including the sheet name 
		 *            and the actual data.
		 * @param specialRows A list to track the rows that require special handling, such as rows for special fields.
		 * 
		 * @return A `SheetBuilder` object representing the created sheet with the populated data.
		 */
		private BuilderExcel.SheetBuilder exGeneralitiesTable(List<SpecialField> specialFieldsDatiGenerali, ExportBaseTable obj, List<Integer> specialRows) {
			BuilderExcel.SheetBuilder sheetBuilder = checkActualSheet(null, obj.getSheet());
			generateTitle(sheetBuilder, obj,true);
			populateSheetVerticalTableKeyValue(builderExcel, obj,sheetBuilder);
			setEmptyRowForSpecialFields(specialFieldsDatiGenerali, specialRows, sheetBuilder);
			return sheetBuilder;
		}

		/**
		 * Sets empty rows for special fields in the general data section of the Excel sheet.
		 * If there are any special fields in the provided general data, empty rows will be added initially.
		 * These rows will be filled with the corresponding values only after the main tables are populated.
		 * 
		 * @param specialFieldsGeneralData A list of special fields that require empty rows. These fields
		 *                                   may need to be handled separately and will have corresponding 
		 *                                   empty rows generated initially.
		 * @param specialRows A list to track the row numbers where special fields have been placed. 
		 *                    These rows will be populated later with their corresponding values.
		 * @param sheetBuilder The builder used to create the sheet and manage rows. It is responsible for 
		 *                     adding the empty rows at the appropriate locations.
		 */
		private void setEmptyRowForSpecialFields(List<SpecialField> specialFieldsGeneralData,List<Integer> specialRows, BuilderExcel.SheetBuilder sheetBuilder) {
			if(specialFieldsGeneralData != null && !specialFieldsGeneralData.isEmpty()) {
				specialFieldsGeneralData.forEach(sf -> {
					specialRows.add(sheetBuilder.rowNum);
					sheetBuilder.createEmptyRow(specialFieldsGeneralData.size());
				});
			}
		}

		
		/**
		 * Generates tables based on the `classes.ComplexExport` class within the context of generating a full report
		 * (`generateReportExcel`). The `classes.ComplexExport` class allows for the setting of extra features such as:
		 * <ul>
		 *     <li>Table reference label</li>
		 *     <li>Special columns (Rows with formulas for handling calculations)</li>
		 *     <li>Pivot summaries</li>
		 *     <li>Specifying the sheet on which the Excel file will be generated</li>
		 * </ul>
		 * 
		 * This method populates the report with the necessary tables, taking into account any special rows 
		 * or special fields defined for the general data.
		 * 
		 * @param sheetBuilder The builder used to construct the sheet in the Excel file.
		 * @param specialRows A list of row indices where special handling (such as formulas or empty rows) 
		 *                    should be applied.
		 * @param specialFieldsDatiGenerali A list of special fields from the general data that may require 
		 *                                  special handling or empty rows.
		 * @param table The `classes.ComplexExport` table object that contains the data and additional settings for
		 *              generating the table. This object allows for customizations like pivot tables, special columns, 
		 *              and a table reference label.
		 * @return The updated `sheetBuilder` containing the populated table.
		 */
		private BuilderExcel.SheetBuilder exComplexExportForReport(BuilderExcel.SheetBuilder sheetBuilder, List<Integer> specialRows, List<SpecialField> specialFieldsDatiGenerali, ComplexExport<? extends ExportSimple> table) {
			sheetBuilder = exGenerateComplexExcel(sheetBuilder, table);
			List<? extends ExportSimple> dati = table.getDati();
			ExportSimple t = dati.get(0);
			for(Integer specialRow : specialRows) {
				BuilderExcel.RowBuilder<?> rowBuilder =  sheetBuilder.createRow(t, sheetBuilder.sheet.getRow(specialRow));
				specialFieldsDatiGenerali.forEach(field -> {
					generateSpecialColumnNoOrder(table.getDati().size(), field, t, rowBuilder);
				});
			}
			cellBeforeCopy = cellBefore;
			return sheetBuilder;
		}

		/**
		 * Method to generate tables based on the `classes.ComplexExport` class.
		 * The `classes.ComplexExport` class allows for the configuration of advanced elements, including:
		 * <ul>
		 *   <li>Table reference label</li>
		 *   <li>Special columns (rows with formulas for handling calculations)</li>
		 *   <li>Pivot tables for summarization</li>
		 *   <li>Specifying the sheet on which to generate the Excel file</li>
		 * </ul>
		 * 
		 * This method takes one or more tables represented by the `classes.ComplexExport` class and
		 * generates an Excel file containing all the tables. Each table can have advanced configurations such as 
		 * special columns and pivot tables. The generated Excel file is returned as a byte array.
		 * 
		 * @param reports A list of tables represented in full using the `classes.ComplexExport` class.
		 *                Each table can have advanced configurations like labels, special columns, and pivots.
		 * @return A byte array (byte[]) representing the generated Excel file, ready to be downloaded or saved.
		 */
		@SafeVarargs
		public final byte[] generateComplexExcel(ComplexExport<? extends ExportSimple>... reports) {

			builderExcel = generateBuilderExcel();
			
			List<ComplexExport<? extends ExportSimple>> reportsList = Arrays.asList(reports);
			
			BuilderExcel.SheetBuilder sheetBuilder = null;
			for(ComplexExport<? extends ExportSimple> report : reportsList){				
				sheetBuilder = exGenerateComplexExcel(sheetBuilder, report);
			};

			if(builderExcel != null)
				builderExcel.autoSizeColumns(getAutoSize());
			
			if(AlignExcel.VERTICAL.equals(orientation))
				sheetBuilder.freezePane(colFreezePane, rowFreezePane);

            assert builderExcel != null;

            return builderExcel.build();
		}

		private boolean header = true;
		
		/**
		 * Method to process each individual `classes.ComplexExport` object.
		 * This method processes a given `classes.ComplexExport` and populates the corresponding Excel sheet.
		 * It checks the current sheet, retrieves special fields, reference labels, and data to be written,
		 * and then populates the sheet with the data, handling any special columns or pivot configurations.
		 *
		 * @param sheetBuilder The builder used to create the Excel sheet.
		 * @param report The `classes.ComplexExport` object representing the table and its associated configuration.
		 * @return The updated `SheetBuilder` after processing the `classes.ComplexExport` and populating the sheet.
		 */
		private BuilderExcel.SheetBuilder exGenerateComplexExcel(BuilderExcel.SheetBuilder sheetBuilder, ComplexExport<? extends ExportSimple> report) {
			sheetBuilder = checkActualSheet(sheetBuilder, report.getSheet());
			List<SpecialField> specialFields = report.getSpecialFields();
			List<ReferenceLabel> referenceLabels = report.getReferenceLabels();
			List<? extends ExportSimple> entitiesToWrite = report.getDati();
			entities.add(entitiesToWrite);
			pivot = report.getPivot();
			header = report.isHeader();
			populateSheetStandardTable(entitiesToWrite, sheetBuilder,specialFields,referenceLabels);
			return sheetBuilder;
		}

		/**
		 * Method to generate an Excel file with simple tables.
		 * 
		 * The objects in the list must extend the `interfaces.ExportSimple` class and must use the annotation system
		 * based on the `@ExcelExport` interface.
		 * 
		 * This method creates an Excel sheet, writes the entities to the sheet as tables, and applies standard 
		 * formatting and settings.
		 * 
		 * @param entitiesToWrite A list of entities to be written into the Excel file in table format. 
		 *                        These entities must implement the `interfaces.ExportSimple` interface and be annotated with `@ExcelExport`.
		 * @return A byte array (`byte[]`) representing the generated Excel file.
		 */
		@SafeVarargs
		public final byte[] generateExcel(List<? extends ExportSimple>... entitiesToWrite) {
			builderExcel = generateBuilderExcel();
			
			final BuilderExcel.SheetBuilder sheetBuilder = builderExcel.createSheet("export");
			Collections.addAll(entities, entitiesToWrite);
			entities.stream().forEach(entity -> {
				populateSheetStandardTable(entity, sheetBuilder,null,null);
			});
			
			sheetBuilder.autoSizeColumn(getAutoSize());
			
			if(AlignExcel.VERTICAL.equals(orientation))
				sheetBuilder.freezePane(colFreezePane, rowFreezePane);
			
			byte[] response = builderExcel.build();
			
			return response;
		}

		/**
		 * Method that processes objects representing a single table and generates the table 
		 * according to the orientation adopted for the Excel file.
		 * 
		 * Depending on the orientation (vertical or horizontal), it will call the appropriate 
		 * method to generate the table in the Excel sheet.
		 * 
		 * @param <T> The type of the entities to be written. It must extend `interfaces.ExportSimple`.
		 * @param entitiesToWrite A list of entities to be written into the Excel sheet.
		 *                        These entities should implement `interfaces.ExportSimple`.
		 * @param sheetBuilder The builder responsible for constructing the sheet.
		 * @param specialFields A list of special fields (e.g., fields with formulas) to be handled in the table.
		 * @param labels A list of reference labels for additional context in the table (e.g., column labels).
		 */
		private <T extends ExportSimple> void populateSheetStandardTable(List<T> entitiesToWrite, BuilderExcel.SheetBuilder sheetBuilder, List<SpecialField> specialFields, List<ReferenceLabel> labels) {
			switch(orientation) {
			case VERTICAL:
				exVerticalTable(sheetBuilder, entitiesToWrite, specialFields,labels);
				break;
			case HORIZONTAL:
				exHorizontalTable(sheetBuilder,entitiesToWrite,specialFields,labels);
				break;
			default:
				throw new RuntimeException("Error while generating excel file");
			}
		}

//================================================================================================================================
		/**
		 * VERTICAL
		 */
//================================================================================================================================
		
		/**
		 * Method for processing a single table with a vertical orientation.
		 * 
		 * This method processes the entities to be written into the sheet in a vertical format,
		 * generating the appropriate headers, values, special columns, and pivot tables if applicable.
		 * 
		 * @param <T> The type of the entities to be written. It must extend `interfaces.ExportSimple`.
		 * @param sheetBuilder The builder responsible for constructing the sheet.
		 * @param entitiesToWrite A list of entities to be written into the sheet in a vertical format.
		 * @param specialFields A list of special fields (e.g., fields with formulas) to be handled in the table.
		 * @param labels A list of reference labels for additional context in the table (e.g., column labels).
		 */
		private <T extends ExportSimple> void exVerticalTable(BuilderExcel.SheetBuilder sheetBuilder, List<T> entitiesToWrite, List<SpecialField> specialFields, List<ReferenceLabel> labels) {
			T t = entitiesToWrite.get(0);
			generateReferenceLabels(sheetBuilder, labels, t);
			generateTitle(sheetBuilder,t,false);
			
			tempStartRow = sheetBuilder.rowNum;
			tempStartGrow = tempStartRow;
			if(header)
				generateHeaders(rowBuilderByOrientation(t, sheetBuilder), t);
			generateColumnsValue(entitiesToWrite, sheetBuilder);
			
			generateSpecialColumns(entitiesToWrite.size(), specialFields, t, sheetBuilder);
			sheetBuilder.createEmptyRow(distanceTable);
			generatePivot(entitiesToWrite, t, sheetBuilder);
			sheetBuilder.createEmptyRow(distanceTable);
		}
		
		/**
		 * Method for generating a row using the values from the objects in the list (orientation = VERTICAL).
		 * This method processes each entity in the provided list and generates a row for each, 
		 * filling the cells with the respective values based on the object's data.
		 * 
		 * @param <T> The type of the entities to be written. It must extend `interfaces.ExportBaseInterface`.
		 * @param entitiesToWrite The list of entities whose values will be used to generate rows in the Excel sheet.
		 * @param sheetBuilder The builder responsible for creating the rows and managing the sheet structure.
		 */
		private <T extends ExportBaseInterface> void generateColumnsValue(List<T> entitiesToWrite, BuilderExcel.SheetBuilder sheetBuilder) {
			entitiesToWrite.stream().forEach(entity -> {
				BuilderExcel.RowBuilder<?> rowBuilder = sheetBuilder.createRow(entity);
				exCol(entity, rowBuilder);
			});
		}
		
//================================================================================================================================
		/**
		 * HORIZONTAL
		 */
//================================================================================================================================
		/**
		 * Temporary variable that indicates the starting row from which the Excel file should be generated.
		 * (USED ONLY IF orientation = HORIZONTAL)
		 */
		private Integer tempStartRow = null;

		/**
		 * Temporary variable that indicates the row we are currently pointing to.
		 * (USED ONLY IF orientation = HORIZONTAL)
		 */
		private int tempStartGrow = 0;

		/**
		 * Temporary variable that indicates if we are generating the first table in the Excel sheet.
		 * (USED ONLY IF orientation = HORIZONTAL)
		 */
		private boolean firstTable = true;

		/**
		 * Method for generating a table with a Horizontal orientation.
		 * 
		 * @param <T> Type of the entities to write into the Excel sheet
		 * @param sheetBuilder The builder object responsible for constructing the Excel sheet
		 * @param entitiesToWrite The list of entities to write into the table
		 * @param specialFields Special fields, such as calculated columns or formulas, for the table
		 * @param labels The reference labels to be applied in the table
		 */
		private <T extends ExportSimple> void exHorizontalTable(BuilderExcel.SheetBuilder sheetBuilder, List<T> entitiesToWrite, List<SpecialField> specialFields, List<ReferenceLabel> labels) {
			T t = entitiesToWrite.get(0);
			generateTitle(sheetBuilder, t,false);
			setTemporaryVariables(sheetBuilder);
			if(header)
				generateHeaders(rowBuilderByOrientation(t, sheetBuilder),t);		
			generateColumnsValueHorizontal(entitiesToWrite, sheetBuilder);
			generateSpecialColumns(entitiesToWrite.size(), specialFields, t, sheetBuilder);
			generateReferenceLabels(sheetBuilder, labels, t);
			generatePivot(entitiesToWrite, t, sheetBuilder);
			cellBefore += (getMaxOrder(entitiesToWrite.get(0)) + 1 + distanceTable);
			firstTable = false;
		}

		/**
		 * Sets temporary variables used for positioning the rows when generating a horizontal table.
		 * These variables help track the starting row for the table and the current row position.
		 *
		 * @param sheetBuilder The builder object responsible for constructing the Excel sheet
		 */
		private void setTemporaryVariables(BuilderExcel.SheetBuilder sheetBuilder) {
			if(tempStartRow == null) {
				tempStartRow = sheetBuilder.rowNum;
				tempStartGrow = tempStartRow;
			} else {
				tempStartGrow = tempStartRow;
			}
		}
		
		/**
		 * Method to generate a row using values from an object (for orientation = HORIZONTAL).
		 * It creates rows in a horizontal orientation, handling the current row position and adjusting it accordingly.
		 *
		 * @param <T> Type of the entities to write to the Excel sheet (must implement interfaces.ExportBaseInterface)
		 * @param entitiesToWrite List of entities to be written to the sheet
		 * @param sheetBuilder The builder object responsible for constructing the Excel sheet
		 */
		private <T extends ExportBaseInterface> void generateColumnsValueHorizontal(List<T> entitiesToWrite, BuilderExcel.SheetBuilder sheetBuilder) {
			entitiesToWrite.stream().forEach(entity -> {
				BuilderExcel.RowBuilder<?> rowBuilder = null;
				if(tempStartRow != null && !firstTable)
					rowBuilder = sheetBuilder.createRow(entity,sheetBuilder.sheet.getRow(tempStartGrow++));
				else
					rowBuilder = sheetBuilder.createRow(entity);

				exCol(entity, rowBuilder);
			});
		}

//================================================================================================================================
		/**
		 * COMMON
		 */
//================================================================================================================================
		
		/**
		 * Method to generate the table title, in case the class is annotated with @TableExcel.
		 * It handles the creation of a title row in the Excel sheet.
		 * 
		 * @param <T> Type of entity being processed
		 * @param sheetBuilder The builder responsible for constructing the Excel sheet
		 * @param t The entity object that holds data for the row
		 * @param keyValue A flag to determine how the title row is generated
		 */
		private <T extends ExportSimple> void generateTitle(BuilderExcel.SheetBuilder sheetBuilder, T t, boolean keyValue) {
			TableExcel annotation = t.getClass().getAnnotation(TableExcel.class);
			if(annotation != null) {
				if(!firstTable)
					tempStartGrow = tempStartRow - 1;
				BuilderExcel.RowBuilder<?> rowBuilder = rowBuilderByOrientation(t, sheetBuilder);
				
				String tableName = annotation.name();
				Integer maxOrder =  !keyValue ? getMaxOrder(t) + cellBefore : 1;
				Integer minOrder = getMinOrder(t) + cellBefore;
				rowBuilder.createCellTitle(tableName,minOrder,maxOrder).build();
				rowBuilder.setRowHeight(ROW_TITLE_HEIGHT);
			}
		}
		
		/**
		 * Method used to generate labels referring to the table. 
		 * This type of field must be set in the classes.ComplexExport object.
		 * 
		 * @param <T> The type of the entity being processed
		 * @param sheetBuilder The builder responsible for constructing the Excel sheet
		 * @param referenceLabels A list of reference labels to be placed in the sheet
		 * @param t The entity instance that holds the data for the row
		 */
		private <T extends ExportSimple> void generateReferenceLabels(BuilderExcel.SheetBuilder sheetBuilder, List<ReferenceLabel> referenceLabels, T t) {
			if(referenceLabels != null && !referenceLabels.isEmpty()) {
				for(ReferenceLabel label : referenceLabels) {
					BuilderExcel.RowBuilder<?> rowBuilder = rowBuilderByOrientation(t, sheetBuilder);
					Integer maxOrder = getMaxOrder(t) + cellBefore;
					Integer minOrder = getMinOrder(t) + cellBefore;
					rowBuilder.createCellReferenceLabel(label,maxOrder,minOrder).build();
				}
			}
		}
		
		/**
		 * Method to generate the header row of a table.
		 * 
		 * @param <T> The type of the entity being processed (must implement interfaces.ExportBaseInterface)
		 * @param headerBuilder The builder responsible for constructing the header row
		 * @param t The entity instance used to retrieve field data for header generation
		 */
		private <T extends ExportBaseInterface> void generateHeaders(BuilderExcel.RowBuilder<?> headerBuilder, T t) {
			Arrays.asList(t.getClass().getDeclaredFields()).forEach(f ->{
				headerBuilder.createCellHeader(f,cellBefore,null).build();
			});
			headerBuilder.setRowHeight(ROW_HEADER_HEIGHT);
		}
		
		/**
		 * Method to check if the table to be generated is on the current sheet or on a different sheet.
		 * 
		 * @param sheetBuilder The builder used to create and manipulate the Excel sheet
		 * @param sheetName The name of the sheet to check against
		 * @return The updated SheetBuilder object, either with the existing sheet or a newly created sheet
		 */
		private BuilderExcel.SheetBuilder checkActualSheet(BuilderExcel.SheetBuilder sheetBuilder, String sheetName) {
			if(sheetBuilder == null) {
				return builderExcel.createSheet(sheetName);
			}
			if(!StringUtils.equals(sheetBuilder.sheet.getSheetName(), sheetName)) {
				setTempParameterOldSheet(sheetBuilder);
				BuilderExcel.SheetBuilder createSheet = builderExcel.createSheet(sheetName);
				setTempParameterNewSheet(createSheet);
				return createSheet;
			}
			return sheetBuilder;
		}

		/**
		 * @param createSheet
		 */
		private void setTempParameterNewSheet(BuilderExcel.SheetBuilder createSheet) {
			tempStartRow = createSheet.tempStartRow;
			tempStartGrow = createSheet.tempStartGrow;
			cellBefore = createSheet.cellBefore;
			firstTable = createSheet.firstTable;
		}

		/**
		 * @param sheetBuilder
		 */
		private void setTempParameterOldSheet(BuilderExcel.SheetBuilder sheetBuilder) {
			sheetBuilder.cellBefore = cellBefore;
			sheetBuilder.tempStartRow = tempStartRow;
			sheetBuilder.tempStartGrow = tempStartGrow;
			sheetBuilder.firstTable = firstTable;
		}
		
		/**
		 * Method to generate a cell for each field in the entity.
		 * 
		 * @param <T> The type of the entity (must extend interfaces.ExportBaseInterface)
		 * @param entity The entity whose fields will be used to generate the row cells
		 * @param rowBuilder The builder used to construct the row in the sheet
		 */
		private <T extends ExportBaseInterface> void exCol(T entity, BuilderExcel.RowBuilder<?> rowBuilder) {
			List<Field> fields = Arrays.asList(entity.getClass().getDeclaredFields());
			fields.forEach(field ->{
				rowBuilder.createCellValue(field,cellBefore).build();
			});
		}
		
		/**
		 * Generates special columns for the given table, applying the specified special field operations to each column.
		 * 
		 * This method processes a list of special fields, which are used to generate custom columns that are not part of the regular table data. 
		 * For each special field, it creates a header and applies the corresponding operation (either a custom operation or a standard operation).
		 * The method handles the special fields based on their order and their associated operations.
		 * 
		 * @param <T> The type of the object being exported, which must extend {@link ExportSimple}.
		 * @param sizeTable The size of the table, used for adjusting row calculations in special field operations.
		 * @param specialFields A list of {@link SpecialField} objects, each representing a special field that will be added to the table.
		 *                      If null or empty, no special fields will be processed.
		 * @param t The object of type T that will be used to generate the rows. This object defines how the data should be laid out.
		 * @param sheetBuilder The {@link BuilderExcel.SheetBuilder} instance responsible for creating the sheet and managing row generation.
		 */
		private <T extends ExportSimple> void generateSpecialColumns(int sizeTable,List<SpecialField> specialFields, T t, BuilderExcel.SheetBuilder sheetBuilder) {
			if(specialFields != null) {
				List<Field> fields = Arrays.asList(t.getClass().getDeclaredFields());
				
				for(SpecialField special : specialFields){
					BuilderExcel.RowBuilder<?> rowBuilder = rowBuilderByOrientation(t, sheetBuilder);
					
					rowBuilder.createCellHeaderSpecial(special.getLabel(), special.getOrder() + cellBefore).build();
					
					OperationEnum operation = special.getOperation();
					if(OperationEnum.CUSTOM.equals(operation)) {
						specialFieldCustomOperation(sizeTable, special, rowBuilder, fields,false);					
					} else {
						specialFieldStandardOperation(sizeTable, special, rowBuilder, fields, operation,false);
					}
				};
			}
		}

		/**
		 * Creates a new row builder for the given object `t` based on the specified orientation (horizontal or vertical).
		 * The method determines the row creation logic depending on whether the orientation is horizontal or vertical.
		 * 
		 * When the orientation is horizontal, the method will either reuse an existing row if `tempStartRow` is set and `firstTable` is false, 
		 * or create a new row from the start. When the orientation is vertical, it simply creates a new row.
		 * 
		 * @param <T> The type of the object being exported, which must extend {@link ExportSimple}.
		 * @param t The object of type T that will be used to generate the row. This object defines how the data should be laid out.
		 * @param sheetBuilder The {@link services.ExcelUtility.BuilderExcel.SheetBuilder} instance responsible for building the sheet and managing row creation.
		 * @return A {@link services.ExcelUtility.BuilderExcel.RowBuilder} instance used for building and populating the row.
		 * @throws RuntimeException If the orientation is neither horizontal nor vertical.
		 */
		private <T extends ExportSimple> BuilderExcel.RowBuilder<?> rowBuilderByOrientation(T t, BuilderExcel.SheetBuilder sheetBuilder) {
			BuilderExcel.RowBuilder<?> rowBuilder = null;
			switch(orientation) {
			case HORIZONTAL:
				if(tempStartRow != null && !firstTable)
					rowBuilder = sheetBuilder.createRow(t,sheetBuilder.sheet.getRow(tempStartGrow++));
				else
					rowBuilder = sheetBuilder.createRow(t);
				break;
			case VERTICAL:
				rowBuilder = sheetBuilder.createRow(t);
				break;
			default:
				throw new RuntimeException("Error while generating excel file");
			}
			return rowBuilder;
		}
		
		/**
		 * Method that performs autosize for columns.
		 * Based on the minimum and maximum order declared within the {@link ExcelUtility} annotations.
		 * 
		 * @return The autosize value (short) for columns.
		 */
		private short getAutoSize() {
			short autoSize = 0;
			switch(orientation) {
			case HORIZONTAL:
				Comparator<ExportSimple> comparator = compareEntitiesByMaxOrder();
				autoSize = entities.stream().flatMap(e -> Stream.of(e.stream().max(comparator).map(ExcelUtility::getMaxOrder).orElse(0)))
													   .map(e -> Integer.sum(e, distanceTable))
													   .reduce(0,Integer::sum)
													   .shortValue();
				break;
			case VERTICAL:
				autoSize = entities.stream().flatMap(List::stream)
											.map(ExcelUtility::getMaxOrder)
											.max(Integer::compareTo)
											.orElse(0)
											.shortValue();
			}
			entities.clear();
			return autoSize;
		}

		/**
		 * Applies a custom formula operation to a specified range of cells based on the configuration in the {@link SpecialField}.
		 * The formula is provided by the {@link SpecialField} and placeholders are replaced with the corresponding cell ranges.
		 * 
		 * This method is typically used when a custom formula needs to be applied to the columns defined in the {@link SpecialField},
		 * with each placeholder in the formula replaced by the corresponding range of cells from the table.
		 * 
		 * @param sizeTable The number of rows in the table, used to determine the cell range for the formula.
		 * @param specialField The {@link SpecialField} object that contains the custom formula, columns to apply the operation to,
		 *                     and the style to apply to the resulting cell.
		 * @param rowBuilder The {@link services.ExcelUtility.BuilderExcel.RowBuilder} used to create and populate the row with data and formulas.
		 * @param fields The list of fields representing the columns in the table, used to retrieve the order and style for each column.
		 * @param keyValue A boolean flag that determines if the cell should be treated as a key-value pair (if `true`).
		 */
		private void specialFieldCustomOperation(int sizeTable, SpecialField specialField, BuilderExcel.RowBuilder<?> rowBuilder, List<Field> fields, boolean keyValue) {
			String operationFormula = null;
			operationFormula = specialField.getFormula();
			String[] placeholders = specialField.getColumns();
			CellStyleEnum style = specialField.getStyle();
			Integer order = null;
			int keyValueOrder = specialField.getOrder();
			for (String p : placeholders) {
				Entry<Integer, CellStyleEnum> os = getFieldOrderAndStyleExport(fields, p);
				if(os != null) {
					if(style == null)
						style = os.getValue();
					if(order == null)
						order = os.getKey();
					
					String cellMin = CellReference.convertNumToColString(os.getKey() + (keyValue ? cellBeforeCopy : cellBefore)) + (tempStartRow + (header ? 2 : 1));
					String cellMax = CellReference.convertNumToColString(os.getKey() + (keyValue ? cellBeforeCopy : cellBefore)) + (tempStartRow + sizeTable + (header ? 1 : 0));
					String replacer = cellMin + ":" + cellMax;
					operationFormula = operationFormula.replaceFirst(PLACEHOLDER_CELL, replacer);
				} else 
					return;
			}
			
			if(StringUtils.isNotBlank(operationFormula) && style != null)
				if(keyValue)
					rowBuilder.createCellValueWithFormula(++keyValueOrder, style, operationFormula).buildFormula();
				else
					rowBuilder.createCellValueWithFormula(order + cellBefore, style, operationFormula).buildFormula();
		}

		/**
		 * Generates and applies a formula to a specified column in a table based on the provided operation.
		 * The formula is calculated for the range of rows determined by the size of the table, 
		 * and the result is written to the appropriate cell in the row.
		 * 
		 * The method works by iterating through the columns specified in the {@link SpecialField} 
		 * and applying the formula based on the {@link OperationEnum} provided.
		 * 
		 * @param sizeTable The number of rows in the table, used to define the range for the formula.
		 * @param specialField The {@link SpecialField} object containing the column(s) to apply the operation to, and the formula to use.
		 * @param rowBuilder The {@link services.ExcelUtility.BuilderExcel.RowBuilder} used to create and populate the row with data and formulas.
		 * @param fields The list of fields, representing the columns in the table, used to retrieve the order and style information.
		 * @param operation The operation to apply (e.g., SUM, AVERAGE, etc.), specified by the {@link OperationEnum}.
		 * @param keyValue A boolean flag that determines whether the cell should be treated as a key-value pair (if `true`).
		 */
		private void specialFieldStandardOperation(int sizeTable, SpecialField specialField, BuilderExcel.RowBuilder<?> rowBuilder, List<Field> fields, OperationEnum operation, boolean keyValue) {
			String operationFormula = null;
			int keyValueOrder = specialField.getOrder();
			for(String column : specialField.getColumns()) {
				Entry<Integer, CellStyleEnum> orderAndStyle = getFieldOrderAndStyleExport(fields, column);
				if(Objects.nonNull(orderAndStyle)) {
					Integer order = orderAndStyle.getKey() + (keyValue ? cellBeforeCopy : cellBefore);
					CellStyleEnum style = orderAndStyle.getValue();
					operationFormula = operationFormula(operation, order,tempStartRow + (header ? 2 : 1),(tempStartRow + sizeTable + (header ? 1 : 0)));
					if(keyValue)
						rowBuilder.createCellValueWithFormula(++keyValueOrder,style,operationFormula).buildFormula();
					else
						rowBuilder.createCellValueWithFormula(order,style,operationFormula).buildFormula();
				}
			}
		}

//================================================================================================================================
		/**
		 * VERTICAL TABLE KEY VALUE
		 */
//================================================================================================================================
	
		/**
		 * Method to populate key-value tables, as seen in tables used to write general data in the Excel file.
		 * 
		 * @param builderExcel  The Excel builder used for creating the Excel file.
		 * @param obj           The object containing data to be written to the Excel sheet.
		 * @param sheetBuilder  The builder used to construct rows and cells in the sheet.
		 */
		private void populateSheetVerticalTableKeyValue(BuilderExcel builderExcel, ExportBaseTable obj, BuilderExcel.SheetBuilder sheetBuilder) {
			List<Field> declaredFields = Arrays.asList(obj.getClass().getDeclaredFields());
			for(Field f : declaredFields) {
				BuilderExcel.RowBuilder<?> rowBuilder = sheetBuilder.createRow(obj);
				rowBuilder.createCellHeader(f,0,null).build();
				rowBuilder.createCellValue(f,1).build();
			};
		}

		/**
		 * Method to generate a special column by creating both a header and a value based on the order of the special field.
		 * 
		 * @param <T>              The type of the object being exported.
		 * @param sizeTable        The number of rows in the table.
		 * @param specialField     The special field to be used for generating the column.
		 * @param t                The object containing the data.
		 * @param rowBuilder       The builder for constructing rows in the Excel sheet.
		 */
		private <T extends ExportSimple> void generateSpecialColumnNoOrder(int sizeTable,SpecialField specialField, T t, BuilderExcel.RowBuilder<?> rowBuilder) {
			if(specialField != null) {
				List<Field> fields = Arrays.asList(t.getClass().getDeclaredFields());
				
				rowBuilder.createCellHeaderSpecial(specialField.getLabel(),specialField.getOrder()).build();
				
				OperationEnum operation = specialField.getOperation();
				if(OperationEnum.CUSTOM.equals(operation)) {
					specialFieldCustomOperation(sizeTable, specialField, rowBuilder, fields,true);					
				} else {
					specialFieldStandardOperation(sizeTable, specialField, rowBuilder, fields, operation,true);
				}
			}
		}

//================================================================================================================================
		/**
		 * PIVOT 
		 */
//================================================================================================================================
	
		private Pivot pivot;
		
		/**
		 * Generates a pivot table summary in the Excel sheet based on the provided table data.
		 * The pivot is created using the provided formula, with dynamic column conditions and calculations.
		 * This method applies the given pivot formula, processes each unique value from the table,
		 * and inserts the calculated values into the appropriate rows and columns of the sheet.
		 * 
		 * @param <T> The type of object being exported, which extends {@link ExportSimple}.
		 * @param table The list of data entities to be used for pivot calculations.
		 * @param t The instance of the object containing the data, used to get field values and configurations.
		 * @param sheetBuilder The builder used to create the Excel sheet and manage row/column creation.
		 * @throws RuntimeException If there are an incorrect number of placeholders in the pivot formula.
		 */
		private <T extends ExportSimple> void generatePivot(List<? extends ExportSimple> table, T t, BuilderExcel.SheetBuilder sheetBuilder) {
			if(pivot != null) {
				String sheet = "";
				if(StringUtils.isNotBlank(pivot.getSheet())) {
					sheet = sheetBuilder.sheet.getSheetName();
					sheetBuilder = checkActualSheet(sheetBuilder, pivot.getSheet());
				}
				
				boolean changeSheet = StringUtils.isNotBlank(sheet);
				
				List<Field> fieldsTotal = Arrays.asList(t.getClass().getDeclaredFields());
				
				List<Field> fields = new ArrayList<>();
				Stream.concat(Stream.of(pivot.getColumnCondition()), Stream.of(pivot.getColumnToCalculate()))
					  .forEach(column -> fields.add(ExcelUtility.getField(fieldsTotal, null, column)));
				
				List<List<?>> values = table.stream().map(entity -> getFieldValueForFunction(entity,pivot.getColumnCondition()))
												     .filter(f -> !f.isEmpty())
												   	 .distinct()
												   	 .collect(Collectors.toList());
				
				int start = 0 + cellBefore, end = fields.size() - 1 + cellBefore;
				
				List<Entry<Integer,CellStyleEnum>> osColumnsCondition = new ArrayList<>();
				for(String col : pivot.getColumnCondition()) {
					osColumnsCondition.add(getFieldOrderAndStyleExport(fields, col));
				}
				
				List<Entry<Integer,CellStyleEnum>> osColumnsCalculate = new ArrayList<>();
				for(String col : pivot.getColumnToCalculate()) {
					osColumnsCalculate.add(getFieldOrderAndStyleExport(fields, col));
				}
				
				if(StringUtils.isNotBlank(pivot.getLabel())) {
					List<Integer> colOrder = osColumnsCalculate.stream().map(c -> c.getKey()).collect(Collectors.toList());
					colOrder.addAll(osColumnsCondition.stream().map(c -> c.getKey()).collect(Collectors.toList()));
					BuilderExcel.RowBuilder<?> rowBuilderTitle = rowBuilderByOrientation(t, sheetBuilder);
					rowBuilderTitle.createCellTitle(pivot.getLabel(), start, end).build();
					rowBuilderTitle.setRowHeight(ROW_TITLE_HEIGHT);
				}
				
				/*
				 * CREAZIONE HEADER PIVOT
				 */
				BuilderExcel.RowBuilder<?> rowBuilderHeader = rowBuilderByOrientation(t,sheetBuilder);
				for(Field field :fields) {rowBuilderHeader.createCellHeader(field, 0, start++ + cellBefore).build();}
				rowBuilderHeader.setRowHeight(ROW_HEADER_HEIGHT);
				start = 0 + cellBefore;
				
				String formula = pivot.getFormula();
				
				int tempStartRowPivot = sheetBuilder.rowNum;
				
				for(List<?> vs : values){
					start = 0 + cellBefore;
					StringBuffer bufferFormula = new StringBuffer(formula).append("(?");
					BuilderExcel.RowBuilder<?> rowBuilder = rowBuilderByOrientation(t, sheetBuilder);
					int i = 0;
					for(Entry<Integer, CellStyleEnum> osColumnCondition : osColumnsCondition) {
						bufferFormula.append(',');
						String cellMinCondition = CellReference.convertNumToColString(osColumnCondition.getKey() + cellBefore) + (tempStartRow + 2);
						String cellMaxCondition = CellReference.convertNumToColString(osColumnCondition.getKey() + cellBefore) + (tempStartRow + table.size() +1);
						String replacerCondition = (changeSheet ? sheet.concat("!") : "") + cellMinCondition + ":" + cellMaxCondition;
						
						String value = null;
						Object v = vs.get(i);
						if(v instanceof String)
							value = new StringBuffer("\"").append(v).append("\"").toString();
						else
							value = v.toString();
						
						bufferFormula.append(replacerCondition).append(',');
						bufferFormula.append(value);
						rowBuilder.createCellValue(v, osColumnCondition.getValue(), start++).build();
						i++;
					}
					bufferFormula.append(')');
					for(Entry<Integer,CellStyleEnum> osColumnCalculate : osColumnsCalculate) {
						String cellMinCalculate = CellReference.convertNumToColString(osColumnCalculate.getKey() + cellBefore) + (tempStartRow + 2);
						String cellMaxCalculate = CellReference.convertNumToColString(osColumnCalculate.getKey() + cellBefore) + (tempStartRow + table.size() +1);
						String replacerCalculate = (changeSheet ? sheet.concat("!") : "")+ cellMinCalculate + ":" + cellMaxCalculate;
						
						String operationFormula = bufferFormula.toString().replaceFirst(PLACEHOLDER_CELL,replacerCalculate);
						
						rowBuilder.createCellValueWithFormula(start++, osColumnCalculate.getValue(), operationFormula).buildFormula();
					}
				}
				
				if(pivot.getSpecialField() != null) {
					SpecialField specialField = pivot.getSpecialField();
					
					BuilderExcel.RowBuilder<?> rowBuilder = rowBuilderByOrientation(t, sheetBuilder);
					rowBuilder.createCellHeaderSpecial(specialField.getLabel(),specialField.getOrder()).build();
					
					String operationFormula = null;
					OperationEnum operation = specialField.getOperation();
					for(String column : specialField.getColumns()) {
						Entry<Integer, CellStyleEnum> orderAndStyle = getFieldOrderAndStyleExport(fields, column);
						if(Objects.nonNull(orderAndStyle)) {
							Integer order = fields.indexOf(fields.stream().filter(f -> f.getName().equals(column)).findFirst().orElse(null));
							CellStyleEnum style = orderAndStyle.getValue();
							operationFormula = operationFormula(operation, order,tempStartRowPivot + 1,(tempStartRowPivot + values.size()));
							rowBuilder.createCellValueWithFormula(order,style,operationFormula).buildFormula();
						}
					}
				}
				pivot = null;
			}
		}

		/**
		 * Retrieves the values of the specified fields from the given entity.
		 * <p>
		 * This method uses reflection to access the fields of the entity and returns their values based on the provided conditions.
		 * If a field is not found or its value is null, an empty string is added to the result.
		 * </p>
		 *
		 * @param <T>        the type of the entity, extending interfaces.ExportSimple
		 * @param entity     the entity from which to retrieve field values
		 * @param conditions the names of the fields to retrieve
		 * @return a list of field values corresponding to the specified conditions; an empty string is added for missing or null fields
		 * @throws RuntimeException if an error occurs during reflection or field access
		 */

		private <T extends ExportSimple> List<?> getFieldValueForFunction(T entity, String...conditions) {
			try {
				List<Field> fields = Arrays.asList(entity.getClass().getDeclaredFields());
				List<Object> response = new ArrayList<>();
				for(String condition : conditions) {
					Field field = fields.stream().filter(f -> StringUtils.equals(f.getName(), condition)).findFirst().orElse(null);
					Optional<?> fieldValue = null;
					if(field != null) {
						field.setAccessible(true);
						fieldValue = getFieldValue(field, entity);
						field.setAccessible(false);
						if(fieldValue.isPresent())
							response.add(fieldValue.get());
						else
							response.add("");
					} else {
						response.add("");
					}
				}
				return response;
			} catch (IllegalArgumentException | IllegalAccessException | SecurityException e) {
				throw new RuntimeException("Error while generating excel file", e);
			}
		}
//================================================================================================================================
		/**
		 * COMMONS BUILD
		 */
//================================================================================================================================
		
		/**
		 * Generates and returns a configured instance of {@link BuilderExcel} for creating an Excel workbook.
		 * The method sets up the necessary styles, fonts, and cell formatting for different sections of the Excel sheet,
		 * including headers, data cells, currency cells, and title cells.
		 * 
		 * @return A {@link BuilderExcel} instance configured with workbook, styles, and fonts.
		 */
		private BuilderExcel generateBuilderExcel() {
			Workbook workBook = new XSSFWorkbook();
			
			Font fontTextHeader = createFontText(workBook,FONT_NAME,HEADER_TEXT_HEIGHT,true);
			Font fontText = createFontText(workBook,FONT_NAME,COMMON_TEXT_HEIGHT,false);
			Font fontTitle = createFontText(workBook, FONT_NAME,TITLE_TEXT_HEIGHT,true);
			
			CellStyle headerStyle = createHeaderStyle(workBook, fontTextHeader, backgroundColorHeader);
			CellStyle commonStyle = createCommonStyle(workBook, fontText);
			CellStyle dataStyle = createDataStyle(workBook, fontText); // Formato di esempio
			CellStyle currencyCellStyle = createCurrencyStyle(workBook, fontText);
			CellStyle titleStyle = createTitleStyle(workBook, fontTitle, null);
			CellStyle percentageStyle = createPercentageStyle(workBook, fontText);
			
			BuilderExcel builderExcel = new BuilderExcel(workBook,FONT_NAME).setCurrencyStyle(currencyCellStyle)
																			.setDataStyle(dataStyle)
																			.setHeaderStyle(headerStyle)
																			.setNormalStyle(commonStyle)
																			.setTitleStyle(titleStyle)
																			.setPercentageStyle(percentageStyle);
			return builderExcel;
		}
	}
	
	/**
	 * Creates and returns a {@link CellStyle} for displaying currency values in an Excel sheet.
	 * The style includes the following properties:
	 * - Right-aligned text
	 * - Font style defined by the provided {@link Font} object
	 * - Currency format with the Euro symbol (e.g., "1,000.00 €")
	 * 
	 * @param workBook The {@link Workbook} instance where the style will be applied.
	 * @param fontText The {@link Font} object to apply to the currency cells.
	 * @return A configured {@link CellStyle} for currency values.
	 */
	static CellStyle createCurrencyStyle(Workbook workBook, Font fontText) {
		CellStyle currencyCellStyle = workBook.createCellStyle();
		setCommonStyle(fontText, currencyCellStyle, HorizontalAlignment.RIGHT);
        CreationHelper creationHelperEuro = workBook.getCreationHelper();
        // Applicare il formato per l'euro
        currencyCellStyle.setDataFormat(creationHelperEuro.createDataFormat().getFormat("#,##0.00 €"));
        return currencyCellStyle;
	}

	/**
	 * Creates and returns a {@link CellStyle} for displaying date values in an Excel sheet.
	 * The style includes the following properties:
	 * - Left-aligned text
	 * - Font style defined by the provided {@link Font} object
	 * - Date format displayed as "dd/MM/yyyy"
	 * 
	 * @param workBook The {@link Workbook} instance where the style will be applied.
	 * @param fontText The {@link Font} object to apply to the date cells.
	 * @return A configured {@link CellStyle} for date values.
	 */
	static CellStyle createDataStyle(Workbook workBook, Font fontText) {
		CellStyle dataStyle = workBook.createCellStyle();
		setCommonStyle(fontText, dataStyle,HorizontalAlignment.LEFT);
		CreationHelper creationHelper = workBook.getCreationHelper();
		DataFormat dateFormat = creationHelper.createDataFormat();
		dataStyle.setDataFormat(dateFormat.getFormat("dd/MM/yyyy"));
		return dataStyle;
	}

	/**
	 * Creates a CellStyle configured for displaying numeric values as percentages in an Excel workbook.
	 * The style applies a percentage format with two decimal places ("0.00%") and aligns the text to the left.
	 *
	 * @param workBook the {@link Workbook} instance where the style will be created.
	 *                 This represents the Excel workbook.
	 * @param fontText the {@link Font} to be applied to the style. It specifies font-related properties
	 *                 like font name, size, and weight.
	 * @return the configured {@link CellStyle} object with percentage formatting and the specified font.
	 */
	static CellStyle createPercentageStyle(Workbook workBook, Font fontText) {
		CellStyle percentStyle = workBook.createCellStyle();
		setCommonStyle(fontText, percentStyle,HorizontalAlignment.LEFT);
		CreationHelper creationHelper = workBook.getCreationHelper();
		DataFormat format = creationHelper.createDataFormat();
	    percentStyle.setDataFormat(format.getFormat("0.00%"));
		return percentStyle;
	}
	
	/**
	 * Creates and returns a {@link CellStyle} for the header of an Excel sheet.
	 * The style includes the following properties:
	 * - Background color set to the provided color
	 * - Solid foreground fill pattern
	 * - Center-aligned text (both horizontally and vertically)
	 * - Font style defined by the provided {@link Font} object
	 * 
	 * @param workBook The {@link Workbook} instance where the style will be applied.
	 * @param fontTextHeader The {@link Font} object to apply to the header cells.
	 * @param colorBackground The {@link Color} to use as the background color for the header.
	 * @return A configured {@link CellStyle} for header cells.
	 */
	static CellStyle createHeaderStyle(Workbook workBook, Font fontTextHeader, Color colorBackground) {
		CellStyle headerStyle = workBook.createCellStyle();
		headerStyle.setFillForegroundColor(colorBackground);
		headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		setCommonStyle(fontTextHeader, headerStyle, HorizontalAlignment.CENTER);
		return headerStyle;
	}

	/**
	 * Creates and returns a {@link CellStyle} for the title of an Excel sheet.
	 * The style includes the following properties:
	 * - Optionally, a background color if the {@code colorBackground} parameter is not {@code null}
	 * - Solid foreground fill pattern if a background color is provided
	 * - Center-aligned text both horizontally and vertically
	 * - Font style defined by the provided {@link Font} object
	 * 
	 * @param workBook The {@link Workbook} instance where the style will be applied.
	 * @param fontTextTitle The {@link Font} object to apply to the title cells.
	 * @param colorBackground The {@link Color} to use as the background color for the title. If {@code null}, no background color is set.
	 * @return A configured {@link CellStyle} for title cells.
	 */
	static CellStyle createTitleStyle(Workbook workBook, Font fontTextTitle, Color colorBackground) {
		CellStyle titleStyle = workBook.createCellStyle();
		if(colorBackground != null) {
			titleStyle.setFillForegroundColor(colorBackground);
			titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		}
		titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		setCommonStyle(fontTextTitle, titleStyle, HorizontalAlignment.CENTER);
		return titleStyle;
	}
	
	/**
	 * Creates and returns a {@link CellStyle} for general use in Excel cells.
	 * The style includes the following properties:
	 * - Left-aligned text
	 * - Font style defined by the provided {@link Font} object
	 * 
	 * @param workBook The {@link Workbook} instance where the style will be applied.
	 * @param fontText The {@link Font} object to apply to the common cells.
	 * @return A configured {@link CellStyle} for general use in Excel cells.
	 */
	static CellStyle createCommonStyle(Workbook workBook, Font fontText) {
		CellStyle commonStyle = workBook.createCellStyle();
		setCommonStyle(fontText, commonStyle,HorizontalAlignment.LEFT);
		return commonStyle;
	}
	
	/**
	 * Creates and returns a {@link CellStyle} for general use in Excel cells, 
	 * with no borders and text wrapping enabled. The style includes:
	 * - Left-aligned text
	 * - Custom font style defined by the provided {@link Font} object
	 * - Text wrapping enabled for cells that contain multiple lines of text
	 * 
	 * @param workBook The {@link Workbook} instance where the style will be applied.
	 * @param fontText The {@link Font} object to apply to the cells.
	 * @return A configured {@link CellStyle} for general use in Excel cells with no borders.
	 */
	static CellStyle createCommonStyleNoBorder(Workbook workBook, Font fontText) {
		CellStyle commonStyle = workBook.createCellStyle();
		commonStyle.setAlignment(HorizontalAlignment.LEFT);
		commonStyle.setFont(fontText);		
		commonStyle.setWrapText(true);
		return commonStyle;
	}

	/**
	 * @param fontText
	 * @param commonStyle
	 */
	private static void setCommonStyle(Font fontText, CellStyle commonStyle,HorizontalAlignment align) {
		commonStyle.setAlignment(align);
		commonStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		commonStyle.setFont(fontText);
		commonStyle.setBorderBottom(BorderStyle.THIN);
		commonStyle.setBorderLeft(BorderStyle.THIN);
		commonStyle.setBorderRight(BorderStyle.THIN);
		commonStyle.setBorderTop(BorderStyle.THIN);
		commonStyle.setWrapText(true);
	}
	
	/**
	 * Creates and returns a {@link Font} object for text styling in Excel cells.
	 * The font can be customized with the following properties:
	 * - Font name (e.g., "Arial", "Calibri")
	 * - Font size (in points)
	 * - Bold setting
	 * 
	 * @param workBook The {@link Workbook} instance where the font will be applied.
	 * @param fontName The name of the font to be applied (e.g., "Arial", "Calibri").
	 * @param height The font size in points (e.g., 10, 12).
	 * @param bold A boolean indicating whether the font should be bold. `true` for bold, `false` for regular.
	 * @return A {@link Font} object configured with the specified font name, size, and bold setting.
	 */
	static Font createFontText(Workbook workBook,String fontName, short height, boolean bold) {
		Font fontText = workBook.createFont();
		fontText.setFontHeightInPoints(height);
		fontText.setFontName(fontName);
		fontText.setBold(bold);
		return fontText;
	}
	
}

