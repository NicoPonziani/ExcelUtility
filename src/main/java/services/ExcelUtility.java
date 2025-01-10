package services;

import classes.ImportBaseDto;
import interfaces.ExcelImportConfig;
import interfaces.ExportSimple;
import org.apache.commons.io.output.ByteArrayOutputStream;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

import java.io.IOException;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.NumberFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.Map.Entry;

import static services.ExportExcel.createCommonStyleNoBorder;
import static services.ExportExcel.createFontText;

public class ExcelUtility {
	
	public static final String DATE_FORMAT = "dd/MM/yyyy";

	/**
	 * Verifies and matches the labels (column titles) in the provided Excel sheet with the expected configuration.
	 * The method iterates through the cells in the first row (or a specified row) and checks if the cell values 
	 * correspond to the expected labels. It also ensures that all required labels are present.
	 * 
	 * @param cellIterator The iterator for the cells in the first row of the Excel sheet, used to extract the column titles.
	 * @param labels A list of expected label configurations. Each configuration corresponds to an expected column title.
	 * @param forceCell A specific cell to use for the first iteration, if provided. This is useful when a cell needs to be manually forced as the first title.
	 * 
	 * @return A map where the key is the column index, and the value is the corresponding {@link ExcelImportConfig} for that column.
	 *         The map represents the matched column titles with their configuration.
	 * 
	 * @throws RuntimeException If any required label is missing or if a label cannot be matched to a column title.
	 * 
	 * @see {@link ExcelImportConfig}
	 */
	static Map<Integer, ExcelImportConfig> checkTitlesExcel(Iterator<Cell> cellIterator, List<ExcelImportConfig> labels, Cell forceCell) {
		Map<Integer,ExcelImportConfig> response = new HashMap<>();
		boolean isFirstIteration = true;
		while(cellIterator.hasNext() || !labels.isEmpty()){
			Cell cell = null;
			if(isFirstIteration  && Objects.nonNull(forceCell)) {
				cell = forceCell;
				isFirstIteration = false;
			} else {
				cell = cellIterator.next();
			}
			final String value = cell.getStringCellValue();
			ExcelImportConfig label = getLabelByColumnName(labels, value);
			
			if(label == null) {
				continue;
			}
			response.put(cell.getColumnIndex(), label);
			labels.removeIf(l -> l.getId().equals(label.getId()));
		}; 
		
		if(!labels.isEmpty()) {
			for(ExcelImportConfig label: labels) {
				if (label.getRequired()) {
					throw new RuntimeException("Missing label " + label.getColumnTitle());
				}
			}
		}
		return response;
	}
	
	/**
	 * Checks that all required fields in the provided response list are populated. 
	 * If any required field is missing (null), a runtime exception is thrown indicating the missing data.
	 * The method iterates over each item in the response list and validates the fields according to the 
	 * `ExcelImportConfig` list, which specifies whether a field is required and the corresponding column.
	 * 
	 * @param <T> The type of the objects in the response list, which must extend {@link ImportBaseDto}.
	 * @param excelImportConfig A list of {@link ExcelImportConfig} objects that define the required fields 
	 *                          and their associated column mappings.
	 * @param response A list of response objects, where each object represents a row of data that has been 
	 *                 parsed from the Excel file.
	 * 
	 * @throws RuntimeException If any required field is missing (null) in the response objects, a runtime 
	 *                          exception is thrown indicating the missing data along with the column and row information.
	 * 
	 * @see {@link ExcelImportConfig}
	 * @see {@link ImportBaseDto}
	 */
	static <T extends ImportBaseDto> void checkRequiredFields(List<ExcelImportConfig> excelImportConfig,List<T> response) {
		for(ExcelImportConfig conf : excelImportConfig) {
			if(conf.getRequired())
				response.forEach(e -> {
					List<Field> fields = Arrays.asList(e.getClass().getDeclaredFields());
					Field field = getField(fields, conf.getColumnOrder(), conf.getColumnTargetField());
					try {
						Object object = field.get(e);
						if(Objects.isNull(object)) 
				        	throw new RuntimeException("Missing required value for column " + conf.getColumnTitle() + " - row: " + e.getNrRow());
					} catch (IllegalArgumentException | IllegalAccessException e1) {
			        	throw new RuntimeException("Missing required value for column " + conf.getColumnTitle() + " - row: " + e.getNrRow());
					}
				});
		}
	}
	
	/**
	 * Retrieves the {@link ExcelImportConfig} label associated with a given column title from the provided configuration list. 
	 * The method searches for a matching column title, performing a case-insensitive check to find the first match 
	 * where the column title contains the provided value.
	 * 
	 * @param searchParamsConf A list of {@link ExcelImportConfig} objects that contains the configuration
	 *                             for each column, including the column title.
	 * @param value The column title (or part of it) to search for within the list of Excel import configurations.
	 * 
	 * @return The {@link ExcelImportConfig} object that matches the given column title, or {@code null} if no match is found.
	 * 
	 * @see {@link ExcelImportConfig}
	 */
	static ExcelImportConfig getLabelByColumnName(List<ExcelImportConfig> searchParamsConf, String value) {
		if(StringUtils.isBlank(value)) return null;

        return searchParamsConf.stream()
								   .filter(config -> value.toLowerCase().contains(config.getColumnTitle().toLowerCase()))
								   .findFirst().orElse(null);
	}
	

	
	/**
	 * Retrieves the {@link Field} from a list of declared fields based on either an alias or an order index.
	 * The method checks if the field is annotated with either {@link ImportExcel} or {@link ExportExcel} annotations 
	 * and uses the alias or order to find the matching field. The search is case-sensitive for the alias and matches 
	 * either the alias values or the field's name. If a field is marked with the special flag in the {@link ImportExcel} annotation, 
	 * it will be excluded from the search.
	 * 
	 * @param declaredFields A list of declared {@link Field} objects that are part of the class.
	 * @param i The order index that can be used to match a field annotated with the {@link ImportExcel} or {@link ExportExcel} annotation.
	 * @param alias The alias name of the field to search for. It may correspond to the alias defined in the annotations or the field's name.
	 * 
	 * @return The {@link Field} that matches the alias or order, or {@code null} if no matching field is found.
	 */
	static Field getField(List<Field> declaredFields, Integer i, String alias) {
		if(StringUtils.isNotBlank(alias)) {
			Optional<Field> opt = declaredFields.parallelStream().filter(f -> {
				ImportExcel annotation = f.getAnnotation(ImportExcel.class);
				if (Objects.isNull(annotation)) {
					ExportExcel annotationExport = f.getAnnotation(ExportExcel.class);
					if(!Objects.isNull(annotationExport)) {
						String label = annotationExport.label();
						Integer order = annotationExport.order();
						return StringUtils.equalsAny(alias, label,f.getName()) || order.equals(i);
					}
					return false;
				}
				List<String> aliasField = Arrays.asList(annotation.alias());
				Integer order = annotation.order();
				boolean special = annotation.special();
				return (aliasField.contains(alias) || f.getName().equals(alias) || order.equals(i)) && !special;
			}).findFirst();
			if(opt.isPresent())
				return opt.get();
		}
		return null;
	}
	
	/**
	 * Retrieves the order and style for exporting a field based on its alias or field name.
	 * This method checks if a field is annotated with {@link ExportExcel} and matches the provided alias or field name.
	 * If a match is found, it returns the order value defined in the annotation along with the style (either directly from 
	 * the {@link ExportExcel} annotation or from a {@link Formula} annotation if the style is set to {@link CellStyleEnum#FORMULA}).
	 * 
	 * @param declaredFields A list of declared {@link Field} objects that are part of the class.
	 * @param alias The alias or field name to search for in the {@link ExportExcel} annotation.
	 * 
	 * @return A {@link Entry} containing the order of the field and its associated {@link CellStyleEnum} style.
	 *         Returns {@code null} if no matching field is found or if no style is associated with the field.
	 */
	static <T> Entry<Integer,CellStyleEnum> getFieldOrderAndStyleExport(List<Field> declaredFields, String alias) {
		if(StringUtils.isNotBlank(alias)) {
			Optional<Field> opt = declaredFields.parallelStream().filter(f -> {
				ExportExcel annotation = f.getAnnotation(ExportExcel.class);
				if (Objects.isNull(annotation)) {
					return false;
				}
				String label = annotation.label();
				String fieldName = f.getName();
				return StringUtils.equalsAny(alias, label,fieldName);
			}).findFirst();
			if(opt.isPresent()) {
				Field field = opt.get();
				ExportExcel annotation = field.getAnnotation(ExportExcel.class);
				CellStyleEnum style = annotation.style();
				if(CellStyleEnum.FORMULA.equals(style)) {
					Formula annotationFormula = field.getAnnotation(Formula.class);
					if(annotationFormula != null)
						style = annotationFormula.style();
				}
				return Map.entry(annotation.order(), style);
			}
		}
		return null;
	}
	
	/**
	 * Searches for a field in the list of declared fields that has the specified alias and is marked as "special"
	 * via the {@link ImportExcel} annotation. This method filters the fields based on the presence of the alias in 
	 * the {@link ImportExcel} annotation's alias list and checks if the field is marked as special using the 
	 * {@link ImportExcel#special()} annotation attribute.
	 * 
	 * @param declaredFields A list of declared {@link Field} objects to search through.
	 * @param alias The alias to match against the field's {@link ImportExcel#alias()} attribute.
	 * 
	 * @return The {@link Field} object that matches the alias and is marked as special, or {@code null} if no such 
	 *         field is found.
	 */
	static Field getFieldSpecial(List<Field> declaredFields, String alias) {
		if(StringUtils.isNotBlank(alias)) {
			Optional<Field> opt = declaredFields.parallelStream().filter(f -> {
				ImportExcel annotation = f.getAnnotation(ImportExcel.class);
				if (Objects.isNull(annotation)) {
					return false;
				}
				List<String> aliasField = Arrays.asList(annotation.alias());
				boolean special = annotation.special();
				return aliasField.contains(alias) && special;
			}).findFirst();
			if(opt.isPresent())
				return opt.get();
		}
		return null;
	}
	
	/**
	 * Sets the value of a specified field in the provided entity, casting and converting the value based on the 
	 * field's type. The method handles different field types such as {@link BigDecimal}, {@link Integer}, {@link Double}, 
	 * {@link String}, and {@link Date}. If the field's type is not one of these, it attempts to cast the value to the 
	 * field's type. If the conversion or setting fails, the method logs an error and returns an appropriate row status.
	 *
	 * @param <T> The type of the value to set in the field.
	 * @param <E> The type of the entity that contains the field.
	 * @param field The field in the entity to set the value for.
	 * @param entity The entity that contains the field to set.
	 * @param value The value to set in the field.
	 * 
	 * @return {@link ExcelRowStatus#VALUES_ROW} if the value was successfully set, or 
	 *         {@link ExcelRowStatus#EMPTY_ROW} if an error occurred during the process.
	 */
	static <T,E> ExcelRowStatus setField(Field field, E entity, T value) {
		field.setAccessible(true);
		try {
			if(field.getType().equals(BigDecimal.class)) 
				field.set(entity, BigDecimal.valueOf(tryCastNumericField(value)));
			else if(field.getType().equals(Integer.class)) 
				field.set(entity, (int)Math.round(tryCastNumericField(value)));
			else if(field.getType().equals(Double.class)) 
				field.set(entity, tryCastNumericField(value));
			else if(field.getType().equals(String.class)) 
				field.set(entity, tryCastStringField(value));
			else if(field.getType().equals(Date.class))
				field.set(entity, tryCastDateField(value));
			else
				field.set(entity, value);
			
			return ExcelRowStatus.VALUES_ROW;
		} catch (IllegalArgumentException | IllegalAccessException | ParseException e) {
			try {
				field.set(entity,field.getType().cast(value));
				return ExcelRowStatus.VALUES_ROW;
			} catch (Exception e1) {
				return ExcelRowStatus.EMPTY_ROW;
			}
		} finally {
			field.setAccessible(false);
		}
	}

	/**
	 * Attempts to cast the given value to a {@link Double}. If the value is already a {@link Double}, it returns it directly.
	 * If the value is a different type (such as a {@link String} or other numeric types), it tries to parse the value 
	 * using a {@link NumberFormat} with the Italian locale.
	 * 
	 * The method supports parsing numeric values that are formatted in Italian number conventions (e.g., using a comma 
	 * for decimal places).
	 *
	 * @param <T> The type of the value to cast. It could be any object type, but typically numeric types or strings.
	 * @param value The value to be cast to a {@link Double}.
	 * 
	 * @return The value as a {@link Double}.
	 * 
	 * @throws ParseException If the value cannot be parsed into a valid numeric value.
	 */
	private static <T> double tryCastNumericField(T value) throws ParseException {
		if(value instanceof Double)
			return (double)value;
		else {
			NumberFormat parser = NumberFormat.getInstance(Locale.ITALIAN);
			Number number = parser.parse(String.valueOf(value));
			return number.doubleValue();
		}
	}
	
	/**
	 * Attempts to cast the given value to a {@link Date}. If the value is already a {@link Date}, it returns it directly.
	 * If the value is a {@link String} or another type, it attempts to parse the value into a {@link Date} using a 
	 * {@link SimpleDateFormat} with a predefined date format.
	 * 
	 * The method supports parsing date strings that match the specified date format pattern.
	 *
	 * @param <T> The type of the value to cast. It could be any object type, but typically a {@link String} or {@link Date}.
	 * @param value The value to be cast into a {@link Date}.
	 * 
	 * @return The value as a {@link Date}.
	 * 
	 * @throws ParseException If the value cannot be parsed into a valid date based on the predefined date format.
	 */
	private static <T> Date tryCastDateField(T value) throws ParseException {
		if(value instanceof Date)
			return (Date)value;
		else {
			SimpleDateFormat parser = new SimpleDateFormat(DATE_FORMAT);
			Date parse = parser.parse(String.valueOf(value));
			return parse;
		}
	}
	
	/**
	 * @param <T>
	 * @param value
	 * @return
	 */
	private static <T> String tryCastStringField(T value) {
		return value instanceof Double ? String.valueOf((double)value) : String.valueOf(value);
	}
	
	/**
	 * Handles the processing of an Excel {@link Cell} by applying the provided {@link CellHandler} function.
	 * If the cell's value cannot be processed and the value is required, an error is logged and a {@link RuntimeException} is thrown.
	 * 
	 * @param <T> The type of the value to be processed. Typically, this could be any object type such as {@link String}, {@link Integer}, etc.
	 * @param cell The {@link Cell} from the Excel sheet to be processed.
	 * @param required A flag indicating whether the value is mandatory. If true, an error will be thrown if the value cannot be processed.
	 * @param handler The {@link CellHandler} function used to process the cell's value.
	 * 
	 * @return The processed value, or {@code null} if the value is not required and processing fails.
	 * 
	 * @throws RuntimeException If the value is required and cannot be processed.
	 */
	static <T> T handleCell(Cell cell, Boolean required, CellHandler<T> handler) {
		try {
	        return handler.handle(cell);
        } catch (Exception e) {
        	if (required) {
	        	throw new RuntimeException("Missing required data for column " + cell.getColumnIndex() + " - " + cell.getRowIndex());
        	}
        }
		return null;
	}
	
	/**
	 * Skips rows in an Excel sheet iterator until the specified row index is reached.
	 * 
	 * @param start The index of the row to stop at. The first row has index 0.
	 * @param iterator The iterator over the rows of the Excel sheet.
	 * 
	 * @return The row at the specified index, or {@code null} if the index is out of range.
	 */
	static Row skipRow(Integer start, Iterator<Row> iterator) {
		for(int i=0; i<start && iterator.hasNext(); i++) {
			Row next = iterator.next();
			if(next.getRowNum() == (start-1)) {
				return next;
			}
		}
		return null;
	}
	
	/**
	 * Skips cells in an Excel row iterator until the specified column index is reached.
	 * 
	 * @param start The index of the cell to stop at. The first column has index 0.
	 * @param iterator The iterator over the cells in the Excel row.
	 * 
	 * @return The cell at the specified column index, or {@code null} if the index is out of range.
	 */
	static Cell skipCell(Integer start, Iterator<Cell> iterator) {
		for(int i=0; i<start && iterator.hasNext(); i++) {
			Cell next = iterator.next();
			if(next.getAddress().getColumn() == (start-1)) {
				return next;
			}
		}
		return null;
	}
	
	/**
	 * @param excelImportConfig
	 */
	static void checkConfigList(List<ExcelImportConfig> excelImportConfig) {
		if(excelImportConfig.isEmpty())
		    throw new IllegalArgumentException("Configuration for import not found");
	}
	
	/**
	 * Retrieves the value of a field from an entity object and returns it wrapped in an {@link Optional}.
	 * Handles various types including {@code Double}, {@code Integer}, {@code String}, {@code BigDecimal}, 
	 * {@code Date}, and {@code Boolean}. For {@code Boolean}, the value is converted to either "SI" (for true) 
	 * or "NO" (for false).
	 *
	 * @param <T> The type of the entity object.
	 * @param field The field to retrieve the value from.
	 * @param entity The object from which the field value should be retrieved.
	 * 
	 * @return An {@link Optional} containing the field value, or an empty {@link Optional} if the field value is {@code null}.
	 * 
	 * @throws IllegalArgumentException If the field is not accessible.
	 * @throws IllegalAccessException If the field is inaccessible or cannot be read.
	 */
	static <T> Optional<?> getFieldValue(Field field,T entity) throws IllegalArgumentException, IllegalAccessException{
		Class<?> type = field.getType();
		
		if(type.equals(Double.class))
			return Optional.ofNullable((Double)field.get(entity));
		
		if(type.equals(Integer.class))
			return Optional.ofNullable((Integer)field.get(entity));
		
		if(type.equals(String.class))
			return Optional.ofNullable((String)field.get(entity));
		
		if(type.equals(BigDecimal.class))
			return Optional.ofNullable((BigDecimal)field.get(entity));
		
		if(type.equals(Date.class))
			return Optional.ofNullable((Date) field.get(entity));
		
		if(type.equals(Boolean.class)) {
			Optional<Boolean> booleanOpt = Optional.ofNullable((Boolean)field.get(entity));
			if(booleanOpt.isPresent()) {
				if(booleanOpt.get())
					return Optional.of("SI");
				else
					return Optional.of("NO");
			}
			return Optional.ofNullable(null);
		}
		return Optional.ofNullable(field.get(entity));
	}
	
	/**
	 * Generates a SUMIF formula for Excel based on the provided column indices, row range, and condition value.
	 * The formula sums the values in the target column where the corresponding values in the condition column
	 * meet the specified condition.
	 * 
	 * <p>The formula follows the structure of an Excel SUMIF formula:
	 * <code>SUMIF(range, criteria, [sum_range])</code>
	 * </p>
	 * 
	 * @param indexColumnToCalc The index of the column to be summed (e.g., 1 for column "A").
	 * @param indexColumnCond The index of the column containing the condition (e.g., 2 for column "B").
	 * @param startRow The starting row number for the range (1-based index).
	 * @param endRow The ending row number for the range (1-based index).
	 * @param value The condition value used in the formula (the criteria for the SUMIF).
	 * 
	 * @return A string representing the SUMIF formula in Excel format.
	 * 
	 * @see CellReference#convertNumToColString(int) for converting column indices to letter representation.
	 */
	static String operationFormulaSumIf(Integer indexColumnToCalc, Integer indexColumnCond, 
											    Integer startRow, Integer endRow,
											    String value) {
	final String sumIf = "SUMIF";
	String indexColumnToCalculate = CellReference.convertNumToColString(indexColumnToCalc);
	String indexColumnCondition = CellReference.convertNumToColString(indexColumnCond);
	String columnToCalculate = new StringBuffer(indexColumnToCalculate).append(startRow).append(':').append(indexColumnToCalculate).append(endRow).toString();
	String columnCondition = new StringBuffer(indexColumnCondition).append(startRow).append(':').append(indexColumnToCalculate).append(endRow).toString();
	return new StringBuffer(sumIf).append('(')
			  					  .append(columnToCalculate).append(',')
			  					  .append(columnCondition).append(',')
			  					  .append("\"" + value + "\"")
				  					  .toString();
	}

	/**
	 * Generates a formula for performing arithmetic operations (SUM, SUBTRACTION, DIVISION) on specified fields in a given row.
	 * The formula is returned in the appropriate Excel formula format based on the operation type.
	 * 
	 * <p>The method builds formulas for operations such as summing, subtracting, or dividing a set of columns for a particular row.</p>
	 * 
	 * @param operation The arithmetic operation to perform. It must be one of the values from {@link OperationEnum}.
	 * @param fields An array of column indices (1-based) that specify the columns to include in the operation.
	 * @param row The row number where the formula will be applied (1-based).
	 * @param plus An integer to adjust the column indices (for cases where you need to shift columns, e.g., offsetting by a specific number of columns).
	 * 
	 * @return A string representing the Excel formula for the specified operation, or {@code null} if the operation is not recognized.
	 * 
	 * @see OperationEnum for the available operations (SUM, SUBTRACTION, DIVISION).
	 */
	private static String operationFormulaRow(OperationEnum operation, int[] fields, int row,int plus) {
		List<String> fieldsDigit = new ArrayList<>();
		StringBuffer buffer = new StringBuffer();
		for(int f : fields) {
			String indexDigit = CellReference.convertNumToColString(f + plus);
			fieldsDigit.add(new StringBuffer(indexDigit).append(row).toString());
		}
		switch(operation) {
		case SUM:
			final String sum = "SUM";
			return buffer.append(sum)
						 .append('(')
						 .append(String.join(",", fieldsDigit))
						 .append(')')
						 .toString();
		case SUBCTRATION:
			for (String field : fieldsDigit) {
				buffer.append(field).append('-');
			}
			return buffer.deleteCharAt(buffer.lastIndexOf("-")).toString();
		case DIVISION:
			for (String field : fieldsDigit) {
				buffer.append(field).append('/');
			}
			return buffer.deleteCharAt(buffer.lastIndexOf("/")).toString();
		default:
			return null;
		}
	}
	
	/**
	 * Generates an Excel formula for performing an arithmetic operation (e.g., SUM) on a specified column range between two rows.
	 * The formula is returned in the appropriate Excel format based on the operation type.
	 * 
	 * <p>This method currently supports the SUM operation, which sums values from the given column and row range.</p>
	 * 
	 * @param operation The arithmetic operation to perform. It must be one of the values from {@link OperationEnum}.
	 * @param field The column index (1-based) where the operation will be applied.
	 * @param startRow The starting row number (1-based) for the range of values.
	 * @param endRow The ending row number (1-based) for the range of values.
	 * 
	 * @return A string representing the Excel formula for the specified operation, or {@code null} if the operation is not recognized.
	 * 
	 * @see OperationEnum for the available operations (currently, only SUM).
	 */
	static String operationFormula(OperationEnum operation, int field,int startRow, int endRow) {
		String indexDigit = CellReference.convertNumToColString(field);
		String start = new StringBuffer(indexDigit).append(startRow).toString();
		String end = new StringBuffer(indexDigit).append(endRow).toString();
		StringBuffer buffer = new StringBuffer();
		switch(operation) {
		case SUM:
			final String sum = "SUM";
			return buffer.append(sum).append('(')
									 .append(start)
									 .append(':')
									 .append(end)
									 .append(')').toString();
		default:
			return null;
		}
	}
	
	/**
	 * Retrieves the maximum order value from the {@link ExportExcel} annotations of the fields of a given entity.
	 * The method searches for the `services.ExportExcel` annotation on each field and returns the highest value of the `order` attribute.
	 * If no annotation is found, it defaults to returning 0.
	 * 
	 * <p>This method is typically used to determine the maximum order of fields in an export configuration.</p>
	 * 
	 * @param entity The entity whose fields are annotated with {@link ExportExcel}.
	 * @param <T> The type of the entity, which must extend {@link ExportSimple}.
	 * 
	 * @return The maximum order value found among the fields' {@link ExportExcel#order()} values. Returns 0 if no annotations are found.
	 * 
	 * @see ExportExcel for the annotation used to specify the export order.
	 */
	static <T extends ExportSimple> Integer getMaxOrder(T entity) {
		return Arrays.asList(entity.getClass().getDeclaredFields()).stream().map(f -> {
			Optional<ExcelUtility.ExportExcel> opt = Optional.ofNullable(f.getAnnotation(ExcelUtility.ExportExcel.class));
            return opt.map(ExportExcel::order).orElse(0);
        }).max(Integer::compareTo).orElse(0);
	}
	
	/**
	 * Retrieves the minimum order value from the {@link ExportExcel} annotations of the fields of a given entity.
	 * The method searches for the `services.ExportExcel` annotation on each field and returns the lowest value of the `order` attribute.
	 * If no annotation is found, it defaults to returning 0.
	 * 
	 * <p>This method is typically used to determine the minimum order of fields in an export configuration.</p>
	 * 
	 * @param entity The entity whose fields are annotated with {@link ExportExcel}.
	 * @param <T> The type of the entity, which must extend {@link ExportSimple}.
	 * 
	 * @return The minimum order value found among the fields' {@link ExportExcel#order()} values. Returns 0 if no annotations are found.
	 * 
	 * @see ExportExcel for the annotation used to specify the export order.
	 */
	static <T extends ExportSimple> Integer getMinOrder(T entity) {
		return Arrays.asList(entity.getClass().getDeclaredFields()).stream().map(f -> {
			Optional<ExcelUtility.ExportExcel> opt = Optional.ofNullable(f.getAnnotation(ExportExcel.class));
            return opt.map(ExportExcel::order).orElse(0);
        }).min(Integer::compareTo).orElse(0);
	}
	
	/**
	 * Returns a comparator that compares two {@link ExportSimple} entities based on their maximum order values.
	 * The method calculates the maximum order value for each entity using the {@link #getMaxOrder(ExportSimple)} method
	 * and compares them in ascending order.
	 *
	 * <p>This comparator can be used to sort a collection of {@link ExportSimple} entities by their maximum order values.</p>
	 *
	 * @return A comparator that compares two entities by their maximum order values, in ascending order.
	 * 
	 * @see #getMaxOrder(ExportSimple) for how the maximum order value is determined.
	 */
	static Comparator<ExportSimple> compareEntitiesByMaxOrder() {
		return (e1,e2) ->{
			Integer maxE1 = getMaxOrder(e1);
			Integer maxE2 = getMaxOrder(e2);
			return Integer.compare(maxE1, maxE2);
		};
	}
	
	//=========================================================================================================
		/**
		 * CLASSI e INTERFACCE
		 */
	//=========================================================================================================
	
	/**
	 * A functional interface used to define a handler that processes an Excel {@link Cell} and returns a result of type {@link T}.
	 * This interface is designed to handle cell values in various ways, such as extracting, transforming, or validating the content
	 * of the cell in an Excel sheet.
	 * <p>
	 * It is particularly useful when you need to provide custom handling logic for different types of cells (e.g., string, numeric, date).
	 * </p>
	 *
	 * <p>This functional interface can be used in lambda expressions or method references.</p>
	 *
	 * @param <T> the type of the result produced after processing the {@link Cell}
	 *
	 * @see Cell for information about Excel cells
	 */
	@FunctionalInterface
	interface CellHandler<T> {
		
		/**
	     * Processes the given Excel {@link Cell} and returns a result of type {@link T}.
	     * 
	     * @param cell the Excel {@link Cell} to process
	     * @return the processed result of type {@link T}
	     * @throws Exception if any error occurs while handling the cell
	     */
		T handle(Cell cell);
	}
	
	/**
	 * Custom annotation used to mark fields in a class that will be imported from an Excel file.
	 * This annotation provides metadata about how the field should be handled during the import process,
	 * including its alias in the Excel file, the order in which it should appear, and whether it is special.
	 * <p>
	 * This annotation is typically used for mapping Excel columns to the fields of an object during the import process.
	 * </p>
	 *
	 * @see ExportExcel for the counterpart annotation used for exporting data to Excel
	 */
	@Retention(RetentionPolicy.RUNTIME)
	@Target(ElementType.FIELD)
	public @interface ImportExcel {

	    /**
	     * The alias or possible column names that the field can have in the Excel file.
	     * Multiple aliases can be provided to handle cases where the same field might have different column names in the Excel sheet.
	     * 
	     * @return an array of strings representing the possible column names or aliases
	     */
	    public String[] alias();

	    /**
	     * The order of the field in the Excel file. This defines the column position of the field during the import process.
	     * It helps to maintain the correct mapping between Excel columns and object fields.
	     * 
	     * @return the order of the field in the Excel sheet
	     */
	    public int order();

	    /**
	     * A flag indicating whether the field is special.
	     * If set to {@code true}, this field may require custom handling or special processing during the import.
	     * Default value is {@code false}.
	     * 
	     * @return {@code true} if the field is special, {@code false} otherwise
	     */
	    public boolean special() default false;
	}

	
	/**
	 * Custom annotation used to mark fields in a class that will be exported to an Excel file.
	 * This annotation provides metadata about how the field should be handled during the export process,
	 * including its label, order, style, and optional color.
	 * <p>
	 * This annotation is typically used for mapping object fields to Excel columns during the export process.
	 * </p>
	 *
	 * @see ImportExcel for the counterpart annotation used for importing data from Excel
	 */
	@Retention(RetentionPolicy.RUNTIME)
	@Target(ElementType.FIELD)
	public @interface ExportExcel {

	    /**
	     * The label for the field that will be used as the column header in the Excel file.
	     * This label is typically used as the column title in the exported Excel sheet.
	     *
	     * @return the label for the field in the Excel sheet
	     */
	    public String label();

	    /**
	     * The order of the field in the Excel file. This defines the column position of the field during the export process.
	     * It helps to maintain the correct mapping between object fields and Excel columns.
	     * 
	     * @return the order of the field in the Excel sheet
	     */
	    public int order();

	    /**
	     * The style to be applied to the column in the Excel file.
	     * The style could include settings for font, color, alignment, etc.
	     *
	     * @return the cell style to be applied to the column
	     */
	    public CellStyleEnum style();

	    /**
	     * The color to be applied to the column header or the cell during the export.
	     * The default color is {@code IndexedColors.AUTOMATIC}, which means no color is applied.
	     * 
	     * @return the color to be applied to the column
	     */
	    public IndexedColors color() default IndexedColors.AUTOMATIC;
	}

	
	/**
	 * Custom annotation used to mark fields in a class that will have formulas applied to them during the export to Excel.
	 * This annotation allows the user to define the formula operation, the fields involved in the formula, and an optional style.
	 * <p>
	 * This annotation is typically used when you need to apply a formula (such as SUM, AVERAGE, etc.) to a field during the export process.
	 * The formula will be automatically computed when exporting the data to Excel.
	 * </p>
	 * 
	 * @see OperationEnum for the available operations that can be applied in the formula
	 * @see CellStyleEnum for the available styles that can be applied to the formula cell
	 */
	@Retention(RetentionPolicy.RUNTIME)
	@Target(ElementType.FIELD)
	public @interface Formula {

	    /**
	     * The operation to be applied in the formula. This defines the kind of operation (e.g., SUM, AVERAGE, etc.) to be used.
	     * The operation determines how the values from the specified fields will be combined in the formula.
	     * 
	     * @return the operation to be applied in the formula
	     */
	    public OperationEnum operation();

	    /**
	     * The fields that will be involved in the formula. This specifies the columns in the Excel sheet
	     * that the formula will use for its calculation. The fields are referenced by their column indices.
	     *
	     * @return the array of field indices that will be involved in the formula
	     */
	    public int[] fields();

	    /**
	     * The style to be applied to the formula cell in the Excel file.
	     * The default style is {@code CellStyleEnum.TEXT}, but it can be customized.
	     * 
	     * @return the cell style to be applied to the formula result
	     */
	    public CellStyleEnum style() default CellStyleEnum.TEXT;
	}

	
	/**
	 * Custom annotation used to mark a class as an Excel table representation.
	 * This annotation allows you to define a table name that can be used during the Excel export process.
	 * <p>
	 * This annotation is typically used when a class is intended to be exported to Excel and you want to associate it with a specific table name.
	 * </p>
	 * 
	 * @see CellStyleEnum for the available styles that can be applied to the Excel cells
	 */
	@Retention(RetentionPolicy.RUNTIME)
	@Target(ElementType.TYPE)
	public @interface TableExcel {

	    /**
	     * The name of the table in the Excel file. This name can be used as a sheet name or to label the table in some way.
	     * 
	     * @return the name of the table, or an empty string if not provided
	     */
	    public String name() default "";
	}
	
	/**
	 * Enum that defines the possible styles that can be applied to a cell in an Excel sheet.
	 * These styles represent different types of data formats for cells.
	 * 
	 * <p>
	 * This enum is used in conjunction with the `services.ExportExcel` annotation to define the style of a cell during Excel export.
	 * </p>
	 */
	public enum CellStyleEnum {

	    /**
	     * Represents a text style for the cell. 
	     * Used for standard text data in the cell.
	     */
	    TEXT,

	    /**
	     * Represents a number style for the cell.
	     * Used for numeric data in the cell.
	     */
	    NUMBER,

	    /**
	     * Represents a currency style for the cell.
	     * Used for monetary values in the cell.
	     */
	    CURRENCY,

	    /**
	     * Represents a date style for the cell.
	     * Used for date values in the cell.
	     */
	    DATE,

	    /**
	     * Represents a percentage style for the cell.
	     * Used for percentage values in the cell.
	     */
	    PERCENTAGE,

	    /**
	     * Represents a formula style for the cell.
	     * Used for cells containing formulas.
	     */
	    FORMULA
	}

	
	/**
	 * Enum that represents the possible statuses of a row during Excel processing.
	 * It defines the states that a row can be in when exporting data to an Excel file.
	 * 
	 * <p>
	 * The row status helps to decide whether a row should be skipped or processed in a certain way.
	 * </p>
	 */
	enum ExcelRowStatus {

	    /**
	     * Represents an empty row, which should be skipped during Excel export.
	     */
	    EMPTY_ROW,

	    /**
	     * Represents a row that contains data and should be processed during Excel export.
	     */
	    VALUES_ROW,

	    /**
	     * Represents a special row, which might require specific handling or formatting.
	     * Such rows are skipped by default.
	     */
	    SPECIAL;

	    /**
	     * A list of row statuses that should be skipped during the export.
	     * This includes `EMPTY_ROW` and `SPECIAL` statuses.
	     */
	    public static final List<ExcelRowStatus> skip = Arrays.asList(EMPTY_ROW, SPECIAL);
	}

	
	/**
	 * Enum that defines the possible alignment options for Excel cells.
	 * This enum is used to control how the contents of a cell should be aligned (either vertically or horizontally).
	 */
	public enum AlignExcel {

	    /**
	     * Represents vertical alignment of the cell's content.
	     */
	    VERTICAL,

	    /**
	     * Represents horizontal alignment of the cell's content.
	     */
	    HORIZONTAL
	}
	
	/**
	 * Enum that defines the possible mathematical operations that can be used in Excel formulas.
	 * These operations can be applied during the export process to calculate values based on other fields.
	 * 
	 * <p>
	 * This enum is used in conjunction with the `Formula` annotation to specify what operation should be applied.
	 * </p>
	 */
	public enum OperationEnum {

	    /**
	     * Represents a SUM operation, which adds up the values of the specified fields.
	     */
	    SUM,

	    /**
	     * Represents a custom operation, which could be any user-defined operation.
	     * The specifics of this operation can vary.
	     */
	    CUSTOM,

	    /**
	     * Represents a subtraction operation, which subtracts the values of the specified fields.
	     */
	    SUBCTRATION,

	    /**
	     * Represents a division operation, which divides the values of the specified fields.
	     */
	    DIVISION
	}

	
	/**
	 * Represents a special field that can be included in the export process, with customizable properties such as label,
	 * columns, operation type, formula, and style.
	 * <p>
	 * This class allows defining special fields in Excel export scenarios, such as summary rows or calculated fields 
	 * that may require formulas or special operations.
	 * </p>
	 */
	public static class SpecialField {
	    /**
	     * The label to be used for this special field. It typically represents a name or description that is displayed
	     * in the Excel export for the field.
	     */
	    private String label;
	    /**
	     * The order in which this special field appears relative to others in the Excel export.
	     * This field helps control the positioning of special fields.
	     */
	    private int order;
	    /**
	     * The columns involved in the special field. This field contains the names of the columns that are used
	     * for calculations or operations in this special field.
	     */
	    private String[] columns;
	    /**
	     * The type of operation to apply to the special field, such as summing, subtracting, or dividing.
	     * This field defines the operation that should be performed on the specified columns.
	     */
	    private OperationEnum operation;
	    /**
	     * A formula associated with the special field. This formula can be used to calculate the value of the field,
	     * based on the operation and columns.
	     */
	    private String formula;
	    /**
	     * The style applied to the special field during the Excel export. This determines how the cell will be formatted.
	     * Examples include text, currency, date, etc.
	     */
	    private CellStyleEnum style;
	    /**
	     * Default constructor for the SpecialField class. It initializes the object without setting any properties.
	     */
	    public SpecialField() {
	        super();
	    }
	    /**
	     * Constructs a SpecialField with the specified label, operation, columns, and order.
	     * 
	     * @param label     the label for this special field
	     * @param operation the operation to perform (e.g., SUM, SUBTRACTION)
	     * @param columns   the columns that this field operates on
	     * @param order     the order in which this special field should appear
	     */
	    public SpecialField(String label, OperationEnum operation, String[] columns, int order) {
	        super();
	        this.label = label;
	        this.operation = operation;
	        this.columns = columns;
	        this.order = order;
	    }
		public String getLabel() {return label;}
		public void setLabel(String label) {this.label = label;}
		public String[] getColumns() {return columns;}
		public void setColumns(String[] columns) {this.columns = columns;}
		public OperationEnum getOperation() {return operation;}
		public void setOperation(OperationEnum operation) {this.operation = operation;}
		public int getOrder() {return order;}
		public void setOrder(int order) {this.order = order;}
		public String getFormula() {return formula;}
		public void setFormula(String formula) {this.formula = formula;}
		public CellStyleEnum getStyle() {return style;}
		public void setStyle(CellStyleEnum style) {this.style = style;}
	}
	
	/**
	 * Represents a reference label, which consists of a value and an optional bold formatting property.
	 * <p>
	 * This class is used to store the value of a label that may be displayed in an Excel export or another
	 * context, and optionally applies bold formatting to the label.
	 * </p>
	 */
	public static class ReferenceLabel {
	    /**
	     * The value of the reference label. This is the actual text or value that represents the label.
	     */
	    private String value;
	    /**
	     * Indicates whether the label should be displayed in bold.
	     * If true, the label will be bold; if false, it will not be bold.
	     */
	    private boolean bold;
	    /**
	     * Constructs a ReferenceLabel with the specified value. The label will not be bold by default.
	     * 
	     * @param value the value of the label
	     */
	    public ReferenceLabel(String value) {
	        this.value = value;
	        this.bold = false;  // Default value for bold is false
	    }
	    
	    /**
	     * Constructs a ReferenceLabel with the specified value and bold property.
	     * 
	     * @param value the value of the label
	     * @param bold  whether the label should be bold or not
	     */
	    public ReferenceLabel(String value, boolean bold) {
	        this.value = value;
	        this.bold = bold;
	    }
	    
		public String getValue() {return value;}
		public boolean isBold() {return bold;}
	}
	
	/**
	 * Represents a Pivot configuration used for calculating values based on certain columns and conditions.
	 * <p>
	 * This class is designed to define a pivot-like calculation, where the formula (by default, "SUMIFS") is used
	 * to aggregate values based on specific columns and conditions. It also allows associating special fields with the pivot.
	 * </p>
	 */
	public static class Pivot {
	    /**
	     * An array of column names to calculate values for. These columns contain the data that will be aggregated.
	     */
	    private String[] columnToCalculate;
	    /**
	     * An array of column names to apply as conditions for the pivot calculation.
	     * Only values in these columns that meet the condition will be included in the aggregation.
	     */
	    private String[] columnCondition;
	    /**
	     * The formula to apply for the pivot calculation. By default, it is set to "SUMIFS".
	     */
	    private String formula;
	    /**
	     * The label for the pivot. This label can be used to identify the pivot in reports or exports.
	     */
	    private String label;
	    /**
	     * The sheet where the pivot will be applied. This helps in identifying the source of data in case there are multiple sheets.
	     */
	    private String sheet;
	    /**
	     * A special field that might be associated with this pivot, containing additional configuration or operations
	     * related to the pivot calculation.
	     */
	    private SpecialField specialField;
	    /**
	     * Constructs a new Pivot with default settings.
	     * The formula is set to "SUMIFS" by default, and all other fields are initialized to null or empty.
	     */
	    public Pivot() {
	        super();
	        formula = "SUMIFS";  // Default formula is "SUMIFS"
	    }
		public String[] getColumnToCalculate() {return columnToCalculate;}
		public void setColumnToCalculate(String... columnToCalculate) {this.columnToCalculate = columnToCalculate;}
		public String[] getColumnCondition() {return columnCondition;}
		public void setColumnCondition(String... columnCondition) {this.columnCondition = columnCondition;}
		public String getFormula() {return formula;}
		@SuppressWarnings("unused")
		private void setFormula(String formula) {this.formula = formula;}
		public String getLabel() {return label;}
		public void setLabel(String label) {this.label = label;}
		public String getSheet() {return sheet;}
		public void setSheet(String sheet) {this.sheet = sheet;}
		public SpecialField getSpecialField() {return specialField;}
		public void setSpecialField(SpecialField specialField) {this.specialField = specialField;}
	}
//====================================================================================================
	/*
	 * Builder Excel
	 */
//====================================================================================================
	
	/**
	 * BuilderExcel is a utility class designed to help build Excel workbooks.
	 * It provides functionalities to create and manage sheets, rows, cells, and apply different styles to them.
	 */
	static class BuilderExcel {
		private Workbook workbook;
		private CellStyle headerStyle;
		private CellStyle dataStyle;
		private CellStyle currencyStyle;
		private CellStyle normalStyle;
		private CellStyle titleStyle;
		private CellStyle percentageStyle;
		
		private Map<Short,CellStyle> dynamicCellStyleCurrency = new HashMap<>();
		private Map<Short,CellStyle> dynamicCellStyleNormal = new HashMap<>();
		private Map<Short,CellStyle> dynamicCellStyleFormula = new HashMap<>();
		
		
		private String fontName;
		
		private List<SheetBuilder> sheets = new ArrayList<>();
		
		/**
	     * Initializes a BuilderExcel with the provided workbook and font name.
	     *
	     * @param workbook the Excel workbook to be used for building
	     * @param fontName the font name to be used in the Excel sheet
	     */
		public BuilderExcel(Workbook workbook,String fontName) {
			super();
			this.workbook = workbook;
			this.fontName = fontName;
		}
		
		/**
	     * Sets the header style to be used for the workbook.
	     *
	     * @param headerStyle the style to be applied to headers
	     * @return the current BuilderExcel instance for chaining
	     */
	    public BuilderExcel setHeaderStyle(CellStyle headerStyle) {
	        this.headerStyle = headerStyle;
	        return this;
	    }

	    /**
	     * Sets the data style to be used for the workbook.
	     *
	     * @param dataStyle the style to be applied to data cells
	     * @return the current BuilderExcel instance for chaining
	     */
	    public BuilderExcel setDataStyle(CellStyle dataStyle) {
	        this.dataStyle = dataStyle;
	        return this;
	    }

	    /**
	     * Sets the currency style to be used for the workbook.
	     *
	     * @param currencyStyle the style to be applied to currency cells
	     * @return the current BuilderExcel instance for chaining
	     */
	    public BuilderExcel setCurrencyStyle(CellStyle currencyStyle) {
	        this.currencyStyle = currencyStyle;
	        return this;
	    }

	    /**
	     * Sets the normal style to be used for the workbook.
	     *
	     * @param normalStyle the style to be applied to normal cells
	     * @return the current BuilderExcel instance for chaining
	     */
	    public BuilderExcel setNormalStyle(CellStyle normalStyle) {
	        this.normalStyle = normalStyle;
	        return this;
	    }

	    /**
	     * Sets the title style to be used for the workbook.
	     *
	     * @param titleStyle the style to be applied to title cells
	     * @return the current BuilderExcel instance for chaining
	     */
	    public BuilderExcel setTitleStyle(CellStyle titleStyle) {
	        this.titleStyle = titleStyle;
	        return this;
	    }

	    /**
	     * Sets the percentage style to be used for the workbook.
	     *
	     * @param percentageStyle the style to be applied to percentage cells
	     * @return the current BuilderExcel instance for chaining
	     */
	    public BuilderExcel setPercentageStyle(CellStyle percentageStyle) {
	        this.percentageStyle = percentageStyle;
	        return this;
	    }

	    /**
	     * Creates a sheet with the given name and returns a SheetBuilder to modify it.
	     *
	     * @param sheetName the name of the sheet to create
	     * @return a SheetBuilder instance to build the sheet
	     */
	    public SheetBuilder createSheet(String sheetName) {
	        if (!sheets.isEmpty()) {
	            Optional<SheetBuilder> opt = sheets.stream().filter(s -> StringUtils.equals(s.sheet.getSheetName(), sheetName)).findFirst();
	            if (opt.isPresent())
	                return opt.get();
	        }
	        SheetBuilder sheet = new SheetBuilder(sheetName, this);
	        sheets.add(sheet);
	        return sheet;
	    }

	    /**
	     * Automatically resizes all columns up to the specified last column number.
	     *
	     * @param lastCellNum the last column number to resize
	     */
	    public void autoSizeColumns(short lastCellNum) {
	        for (SheetBuilder sheet : sheets)
	            sheet.autoSizeColumn(lastCellNum);
	    }

	    /**
	     * Builds and writes the workbook to a byte array.
	     *
	     * @return the byte array representing the Excel workbook
	     */
	    public byte[] build() {
	        try (ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream()) {
	            workbook.write(byteArrayOutputStream);
	            return byteArrayOutputStream.toByteArray();
	        } catch (IOException e) {
	            throw new RuntimeException(e.getMessage(), e);
	        } finally {
	            try {
	                if (workbook != null)
	                    workbook.close();
	            } catch (IOException e) {
	                throw new RuntimeException(e.getMessage(), e);
	            }
	        }
	    }
		
		//====================================================================================================
			/*
			 * Sheet Builder
			 */
		//====================================================================================================
	
	    /**
	     * SheetBuilder is responsible for building and managing individual sheets in an Excel workbook.
	     * It allows for the creation of rows, columns, and the manipulation of sheet properties such as auto-sizing and freezing panes.
	     */
	    class SheetBuilder {
	        
	        // Parent BuilderExcel instance (not used in the current class).
	        @SuppressWarnings("unused")
	        private BuilderExcel builderFather;

	        // The sheet being built.
	        Sheet sheet;

	        // Row number to keep track of rows being added.
	        int rowNum;

	        // Temporary variables used for tracking row and column placement (not fully utilized in this class).
	        Integer tempStartRow;
	        int tempStartGrow;
	        int cellBefore;
	        
	        boolean firstTable = true;

	        /**
	         * Initializes a SheetBuilder with the specified sheet name.
	         * 
	         * @param sheetName the name of the sheet to create
	         * @param builder   the parent BuilderExcel instance to which this sheet belongs
	         */
	        public SheetBuilder(String sheetName, BuilderExcel builder) {
	            this.sheet = workbook.createSheet(sheetName);
	            this.builderFather = builder;
	            this.rowNum = 0;
	        }

	        /**
	         * Creates a new row for the sheet based on the provided entity.
	         * 
	         * @param entity the entity whose data will be used to populate the row
	         * @param <T>    the type of the entity
	         * @return a RowBuilder to build the row for the provided entity
	         */
	        public <T> RowBuilder<?> createRow(T entity) {
	            RowBuilder<?> rowBuilder = new RowBuilder<>(rowNum, this, entity);
	            rowNum++;
	            return rowBuilder;
	        }

	        /**
	         * Creates a new row for the sheet based on the provided entity and an existing row.
	         * 
	         * @param entity the entity whose data will be used to populate the row
	         * @param row    an existing row that will be used as a template (if null, a new row is created)
	         * @param <T>    the type of the entity
	         * @return a RowBuilder to build the row for the provided entity
	         */
	        public <T> RowBuilder<?> createRow(T entity, Row row) {
	            if (row == null) {
	                return createRow(entity);
	            }
	            RowBuilder<?> rowBuilder = new RowBuilder<>(row, this, entity);
	            return rowBuilder;
	        }

	        /**
	         * Automatically resizes all columns in the sheet based on the contents of the first row.
	         */
	        public void autoSizeColumn() {
	            short lastCellNum = sheet.getRow(0).getLastCellNum();
	            while (lastCellNum >= 0) {
	                sheet.autoSizeColumn(lastCellNum, true);
	                lastCellNum--;
	            }
	        }

	        /**
	         * Automatically resizes columns in the sheet up to the specified last column index.
	         * 
	         * @param lastCellNum the last column index to resize (if less than or equal to 0, all columns are resized)
	         */
	        public void autoSizeColumn(short lastCellNum) {
	            if (lastCellNum <= 0) {
	                autoSizeColumn();
	                return;
	            }
	            while (lastCellNum >= 0) {
	                sheet.autoSizeColumn(lastCellNum, true);
	                lastCellNum--;
	            }
	        }

	        /**
	         * Freezes the panes of the sheet at the specified column and row index.
	         * 
	         * @param col the column index where the freeze will occur
	         * @param row the row index where the freeze will occur
	         */
	        public void freezePane(int col, int row) {
	            sheet.createFreezePane(col, row);
	        }

	        /**
	         * Creates empty rows between tables or sections.
	         * 
	         * @param distanceTable the number of empty rows to create
	         */
	        public void createEmptyRow(int distanceTable) {
	            while (distanceTable > 0) {
	                sheet.createRow(rowNum++);
	                distanceTable--;
	            }
	        }
	    }

		
		//====================================================================================================
			/*
			 * Row Builder
			 */
		//====================================================================================================
	
	    /**
	     * RowBuilder is responsible for constructing individual rows in an Excel sheet.
	     * It allows for creating cells based on field annotations, applying formulas, and managing row heights.
	     * It works in conjunction with the `SheetBuilder` class to populate rows with data.
	     *
	     * @param <T> the type of the entity being processed to populate the row
	     */
	    class RowBuilder<T> {
	        
	        // The row being built.
	        private Row row;

	        // The SheetBuilder that owns this row.
	        private SheetBuilder sheetBuilder;

	        // The entity whose fields will be used to populate the row.
	        private T entity;

	        /**
	         * Constructs a RowBuilder using an existing row, SheetBuilder, and entity.
	         * 
	         * @param row the existing row to be populated
	         * @param sheetBuilder the SheetBuilder responsible for the sheet
	         * @param entity the entity whose fields will populate the row
	         */
	        public RowBuilder(Row row, SheetBuilder sheetBuilder, T entity) {
	            this.row = row;
	            this.sheetBuilder = sheetBuilder;
	            this.entity = entity;
	        }

	        /**
	         * Constructs a RowBuilder by creating a new row at the specified row number.
	         * 
	         * @param rowNum the row number for the new row
	         * @param sheetBuilder the SheetBuilder responsible for the sheet
	         * @param entity the entity whose fields will populate the row
	         */
	        public RowBuilder(int rowNum, SheetBuilder sheetBuilder, T entity) {
	            this.sheetBuilder = sheetBuilder;
	            this.row = sheetBuilder.sheet.createRow(rowNum);
	            this.entity = entity;
	        }

	        /**
	         * Creates a horizontal cell for the specified field in the entity.
	         * 
	         * @param field the field to be placed in the cell
	         * @return a CellBuilder to configure the cell
	         */
	        public CellBuilder createCellValue(Field field) {
	            return cellValueExecute(field, 0);
	        }

	        /**
	         * Creates a horizontal cell for the specified field with a specified number of columns before the cell.
	         * 
	         * @param field the field to be placed in the cell
	         * @param nrColumnsBefore the number of columns before the cell
	         * @return a CellBuilder to configure the cell
	         */
	        public CellBuilder createCellValue(Field field, int nrColumnsBefore) {
	            return cellValueExecute(field, nrColumnsBefore);
	        }

	        /**
	         * Creates a cell with a specified value, style, and order.
	         * 
	         * @param value the value to be placed in the cell
	         * @param style the style of the cell
	         * @param order the order (position) of the cell
	         * @return a CellBuilder to configure the cell
	         */
	        public CellBuilder createCellValue(Object value, CellStyleEnum style, int order) {
	            return new CellBuilder(this, order, style, Optional.of(value));
	        }

	        /**
	         * Executes the logic to create a cell value for the given field with an optional offset for column order.
	         * 
	         * @param field the field to be placed in the cell
	         * @param plus the offset to add to the column order
	         * @return a CellBuilder to configure the cell
	         */
	        private CellBuilder cellValueExecute(Field field, int plus) {
	            field.setAccessible(true);
	            int order = -1;
	            try {
	                ExportExcel annotation = field.getAnnotation(ExportExcel.class);
	                if (annotation != null) {
	                    CellStyleEnum style = annotation.style();
	                    order = annotation.order() + plus;
	                    IndexedColors color = annotation.color();
	                    if (CellStyleEnum.FORMULA.equals(style)) {
	                        return formulaCellValue(field, color, style, order, plus);
	                    }
	                    Optional<?> value = getFieldValue(field, entity);

	                    CellBuilder cellBuilder = new CellBuilder(this, order, style, value);

	                    if (!IndexedColors.AUTOMATIC.equals(color))
	                        cellBuilder.setColor(color);

	                    return cellBuilder;
	                }
	            } catch (IllegalAccessException e) {
                    throw new RuntimeException(e);
                } finally {
	                field.setAccessible(false);
	            }
	            return new CellBuilder(this, order, null, Optional.empty());
	        }

	        /**
	         * Handles the creation of a cell with a formula, based on the field's Formula annotation.
	         * 
	         * @param field the field with the formula annotation
	         * @param color the color for the cell
	         * @param style the style for the cell
	         * @param order the order (position) of the cell
	         * @param plus the offset to add to the column order
	         * @return a CellBuilder to configure the formula cell
	         * @throws IllegalAccessException if the field is not accessible
	         */
	        private CellBuilder formulaCellValue(Field field, IndexedColors color, CellStyleEnum style, Integer order, int plus) throws IllegalAccessException {
	            Formula formulaAnnotation = field.getAnnotation(Formula.class);
	            if (formulaAnnotation != null) {
	                int[] fields = formulaAnnotation.fields();
	                OperationEnum operation = formulaAnnotation.operation();
	                CellStyleEnum styleCell = formulaAnnotation.style();
	                String formula = operationFormulaRow(operation, fields, this.row.getRowNum() + 1, plus);

	                CellBuilder cellBuilder = new CellBuilder(this, order, styleCell, formula);

	                if (!IndexedColors.AUTOMATIC.equals(color))
	                    cellBuilder.setColor(color);

	                return cellBuilder;
	            }

	            return new CellBuilder(this, order, CellStyleEnum.TEXT, Optional.empty());
	        }

	        /**
	         * Creates a cell for the header, based on the field annotation.
	         * 
	         * @param field the field to be placed in the header cell
	         * @param nrColumnsBefore the number of columns before the header cell
	         * @param orderParam the order to place the header cell (optional)
	         * @return a CellBuilder to configure the header cell
	         */
	        public CellBuilder createCellHeader(Field field, int nrColumnsBefore, Integer orderParam) {
	            field.setAccessible(true);
	            int order = -1;
	            try {
	                ExportExcel annotation = field.getAnnotation(ExportExcel.class);
	                if (annotation != null) {
	                    String label = annotation.label();
	                    order = annotation.order() + nrColumnsBefore;
	                    CellBuilder cellBuilder = new CellBuilder(this, Objects.isNull(orderParam) ? order : orderParam, CellStyleEnum.TEXT, Optional.of(label));
	                    cellBuilder.setHeader(true);
	                    return cellBuilder;
	                }
	            } finally {
	                field.setAccessible(false);
	            }
	            return new CellBuilder(this, order, null, Optional.empty());
	        }

	        /**
	         * Creates a special header cell with a label and order.
	         * 
	         * @param label the label for the header cell
	         * @param order the order (position) of the header cell
	         * @return a CellBuilder to configure the special header cell
	         */
	        public CellBuilder createCellHeaderSpecial(String label, int order) {
	            CellBuilder cellBuilder = new CellBuilder(this, order, CellStyleEnum.TEXT, Optional.of(label));
	            cellBuilder.setHeader(true);
	            return cellBuilder;
	        }

	        /**
	         * Sets the height of the row.
	         * 
	         * @param points the height of the row in points
	         */
	        public void setRowHeight(float points) {
	            row.setHeightInPoints(points);
	        }

	        /**
	         * Creates a cell with a formula, based on the provided parameters.
	         * 
	         * @param order the order (position) of the cell
	         * @param style the style of the cell
	         * @param operationFormula the formula for the cell
	         * @return a CellBuilder to configure the formula cell
	         */
	        public CellBuilder createCellValueWithFormula(Integer order, CellStyleEnum style, String operationFormula) {
	            return new CellBuilder(this, order, style, operationFormula);
	        }

	        /**
	         * Creates a cell with a table title, merged across the specified range of columns.
	         * 
	         * @param tableName the name of the table
	         * @param minOrder the minimum column order (start of the range)
	         * @param maxOrder the maximum column order (end of the range)
	         * @return a CellBuilder to configure the title cell
	         */
	        public CellBuilder createCellTitle(String tableName, Integer minOrder, Integer maxOrder) {
	            sheetBuilder.sheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), minOrder, maxOrder));
	            CellBuilder cellBuilder = new CellBuilder(this, minOrder, CellStyleEnum.TEXT, Optional.of(tableName));
	            cellBuilder.setTitle(true);
	            for (int i = minOrder + 1; i <= maxOrder; i++) {
	                Cell cell = row.createCell(i);
	                cell.setCellStyle(normalStyle);
	            }
	            return cellBuilder;
	        }

	        /**
	         * Creates a cell with a reference label, merged across the specified range of columns.
	         * 
	         * @param label the reference label for the cell
	         * @param maxOrder the maximum column order (end of the range)
	         * @param minOrder the minimum column order (start of the range)
	         * @return a CellBuilder to configure the reference label cell
	         */
	        public CellBuilder createCellReferenceLabel(ReferenceLabel label, Integer maxOrder, Integer minOrder) {
	            sheetBuilder.sheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), minOrder, maxOrder));
	            CellBuilder cellBuilder = new CellBuilder(this, minOrder, CellStyleEnum.TEXT, Optional.ofNullable(label));
	            cellBuilder.setReferenceLabel(true);
	            for (int i = minOrder + 1; i <= maxOrder; i++) {
	                row.createCell(i);
	            }
	            return cellBuilder;
	        }
	    }

		
		//====================================================================================================
			/*
			 * Cell Builder
			 */
		//====================================================================================================
	
	    /**
	     * The {@code CellBuilder} class is responsible for building and styling individual cells
	     * in an Excel sheet using Apache POI. It provides flexible configuration for various types of cells, 
	     * such as text, numeric, date, currency, formula, etc., and handles different styles and formulas
	     * dynamically.
	     * 
	     * <p>This class uses the builder pattern to allow for an easy and fluent construction of cells with 
	     * specific styles, values, and formulas.</p>
	     */
		class CellBuilder {
			private RowBuilder<?> rowBuilder;
			private Cell cell;
			private Optional<?> value;
			private int order;
			private CellStyleEnum style;
			private IndexedColors color;
			private String formula;
			private boolean header;
			private boolean title;
			private boolean referenceLabel;
				
			 /**
		     * Constructs a {@code CellBuilder} for a cell with the given order and style.
		     * 
		     * @param rowBuilder the {@code RowBuilder} that contains the row this cell belongs to
		     * @param order the index of the cell in the row
		     * @param style the style to be applied to the cell
		     */
			public CellBuilder(RowBuilder<?> rowBuilder, int order, CellStyleEnum style) {
				this.order = order;
				this.style = style;
				this.rowBuilder = rowBuilder;
			}
			
			/**
		     * Constructs a {@code CellBuilder} for a cell with a given formula.
		     * 
		     * @param rowBuilder the {@code RowBuilder} that contains the row this cell belongs to
		     * @param order the index of the cell in the row
		     * @param style the style to be applied to the cell
		     * @param formula the formula to be applied to the cell
		     */
			public CellBuilder( RowBuilder<?> rowBuilder,int order,CellStyleEnum style,String formula) {
				this.order = order;
				this.style = style;
				this.rowBuilder = rowBuilder;
				this.formula = formula;
			}
			
			/**
		     * Constructs a {@code CellBuilder} for a cell with a given value.
		     * 
		     * @param rowBuilder the {@code RowBuilder} that contains the row this cell belongs to
		     * @param order the index of the cell in the row
		     * @param style the style to be applied to the cell
		     * @param value the value to be set in the cell
		     */
			public CellBuilder( RowBuilder<?> rowBuilder,int order,CellStyleEnum style, Optional<?> value) {
				this.value = value;
				this.order = order;
				this.style = style;
				this.rowBuilder = rowBuilder;
			}
			
			/**
		     * Sets whether this cell is a title cell.
		     * 
		     * @param title {@code true} if this cell is a title; {@code false} otherwise
		     */
		    public void setTitle(boolean title) {
		        this.title = title;
		    }

		    /**
		     * Sets whether this cell is a reference label.
		     * 
		     * @param label {@code true} if this cell is a reference label; {@code false} otherwise
		     */
		    public void setReferenceLabel(boolean label) {
		        this.referenceLabel = label;
		    }

		    /**
		     * Sets whether this cell is a header cell.
		     * 
		     * @param header {@code true} if this cell is a header; {@code false} otherwise
		     */
		    public void setHeader(boolean header) {
		        this.header = header;
		    }

		    /**
		     * Sets the background color of the cell.
		     * 
		     * @param color the {@code IndexedColors} color to set as the background of the cell
		     */
		    public void setColor(IndexedColors color) {
		        this.color = color;
		    }
			
		    /**
		     * Builds the formula cell by applying the appropriate formula and style.
		     * 
		     */
			public void buildFormula() {
				if(StringUtils.isBlank(formula)) {
					cell = rowBuilder.row.createCell(order, CellType.BLANK);
					cell.setCellStyle(normalStyle);
					cell.setBlank();
					return;
				}
				switch(style) {
				case CURRENCY:
					cell = rowBuilder.row.createCell(order,CellType.NUMERIC);
					CellStyle currencyStyleFormula = dynamicFormulaStyle(dynamicCellStyleFormula,currencyStyle,(short)65);						
					cell.setCellStyle(dynamicStyle(dynamicCellStyleFormula, currencyStyleFormula));
					break;
				case DATE:
					cell = rowBuilder.row.createCell(order, CellType.NUMERIC);
					cell.setCellStyle(dataStyle);
					break;
				case NUMBER:
					cell = rowBuilder.row.createCell(order, CellType.NUMERIC);
					CellStyle numericStyleFormula = dynamicFormulaStyle(dynamicCellStyleFormula,normalStyle,(short)66);	
					cell.setCellStyle(dynamicStyle(dynamicCellStyleFormula, numericStyleFormula));
					break;
				case PERCENTAGE:
					cell = rowBuilder.row.createCell(order, CellType.NUMERIC);
					CellStyle stylePercentage = dynamicFormulaStyle(dynamicCellStyleFormula,percentageStyle,(short)67);
					cell.setCellStyle(dynamicStyle(dynamicCellStyleNormal,stylePercentage));
					break;
				case TEXT,FORMULA:
					cell = rowBuilder.row.createCell(order, CellType.STRING);
					cell.setCellStyle(normalStyle);
					break;
				default:
					cell = rowBuilder.row.createCell(order, CellType.STRING);
					cell.setCellStyle(normalStyle);
				}
				
				cell.setCellFormula(formula);
			}

			/**
		     * Builds the cell, applying the appropriate style, value, and formula as needed.
		     * This method determines the cell's content and style based on the provided configuration 
		     * and then creates the cell in the row.
		     */
			public void build(){
				if(order == -1)
					return;
				
				if(StringUtils.isNotBlank(formula)) {
					buildFormula();
					return;
				}
				
				if(value.isEmpty()) {
					cell = rowBuilder.row.createCell(order, CellType.BLANK);
					CellStyle styleEmpty = dynamicStyle(dynamicCellStyleNormal,normalStyle);
					cell.setCellStyle(styleEmpty);
					cell.setBlank();
					return;
				}
				
				if(referenceLabel) {
					ReferenceLabel label = (ReferenceLabel) value.get();
					cell = rowBuilder.row.createCell(order, CellType.STRING);
					cell.setCellStyle(createCommonStyleNoBorder(workbook,createFontText(workbook, fontName, services.ExportExcel.ExportBuilder.COMMON_TEXT_HEIGHT, label.isBold())));
					cell.setCellValue(label.getValue());
					return;
				}
				
				if(title) {
					cell = rowBuilder.row.createCell(order, CellType.STRING);
					cell.setCellStyle(titleStyle);
					cell.setCellValue((String) value.get());
					return;
				}
				
				if (header) {
					cell = rowBuilder.row.createCell(order, CellType.STRING);
					cell.setCellStyle(headerStyle);
					cell.setCellValue((String) value.get());
				} else {
					switch(style) {
					case CURRENCY:
						cell = rowBuilder.row.createCell(order,CellType.NUMERIC);
						CellStyle styleCurrency = dynamicStyle(dynamicCellStyleCurrency,currencyStyle);
						cell.setCellStyle(styleCurrency);
						cell.setCellValue(((BigDecimal)value.get()).doubleValue());
						break;
					case DATE:
						cell = rowBuilder.row.createCell(order, CellType.NUMERIC);
						CellStyle styleData = dynamicStyle(dynamicCellStyleNormal,dataStyle);
						cell.setCellStyle(styleData);
						cell.setCellValue(((Date)value.get()));
						break;
					case NUMBER:
						cell = rowBuilder.row.createCell(order, CellType.NUMERIC);
						CellStyle numericStyle = dynamicStyle(dynamicCellStyleNormal,normalStyle);
						cell.setCellStyle(numericStyle);
						setNumericValue();
						break;
					case PERCENTAGE:
						cell = rowBuilder.row.createCell(order,CellType.NUMERIC);
						CellStyle stylePercentage = dynamicStyle(dynamicCellStyleNormal,percentageStyle);
						cell.setCellStyle(stylePercentage);
						setNumericValue();
						break;
					case TEXT:
						cell = rowBuilder.row.createCell(order, CellType.STRING);
						CellStyle styleText = dynamicStyle(dynamicCellStyleNormal,normalStyle);
						cell.setCellStyle(styleText);
						cell.setCellValue((String)value.get());
						break;
					default:
						cell = rowBuilder.row.createCell(order, CellType.STRING);
						CellStyle styleDefault = dynamicStyle(dynamicCellStyleNormal,normalStyle);
						cell.setCellStyle(styleDefault);
						cell.setCellValue((String)value.get());
					}
				}
			}

			/**
			 * 
			 */
			private void setNumericValue() {
				if(value.get() instanceof Double)
					cell.setCellValue((Double)value.get());
				if(value.get() instanceof Integer)
					cell.setCellValue((Integer)value.get());
				if(value.get() instanceof BigDecimal)
					cell.setCellValue(((BigDecimal)value.get()).doubleValue());
			}

			/**
		     * Applies dynamic styling to the cell based on the given color and base style.
		     * 
		     * @param dynamicCellStyle the dynamic cell style map
		     * @param styleToCopy the base style to copy
		     * @return the final styled cell
		     */
			private CellStyle dynamicStyle(Map<Short,CellStyle> dynamicCellStyle,CellStyle styleToCopy) {
				CellStyle style = styleToCopy;
				if(color != null)
					try {
						style = dynamicCellStyle.get(color.index);
						if(style == null)
							throw new RuntimeException();
					} catch (Exception e) {
						style = workbook.createCellStyle();
						style.cloneStyleFrom(styleToCopy);
						style.setFillForegroundColor(color.getIndex());
						style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					}

				return style;
			}
			
			/**
		     * Applies dynamic styling to formula cells based on the given style and index.
		     * 
		     * @param dynamicCellStyle the dynamic cell style map
		     * @param styleToCopy the base style to copy
		     * @param index the style index
		     * @return the final styled cell for formula cells
		     */
			private CellStyle dynamicFormulaStyle(Map<Short,CellStyle> dynamicCellStyle,CellStyle styleToCopy,short index) {
				CellStyle currencyStyleFormula = styleToCopy;
				try {
					currencyStyleFormula = dynamicCellStyle.get(index);
					if(currencyStyleFormula == null)
						throw new RuntimeException();
				} catch (Exception e) {
					currencyStyleFormula = workbook.createCellStyle();
					currencyStyleFormula.cloneStyleFrom(styleToCopy);
					currencyStyleFormula.setFont(createFontText(workbook, fontName, services.ExportExcel.ExportBuilder.COMMON_TEXT_HEIGHT, true));
					dynamicCellStyleFormula.put(index, currencyStyleFormula);
				}
				return currencyStyleFormula;
			}
			
		}
	}
}
