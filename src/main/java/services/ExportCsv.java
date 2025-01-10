package services;

import java.io.BufferedWriter;
import java.io.ByteArrayOutputStream;
import java.io.OutputStreamWriter;
import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.Comparator;
import java.util.List;
import java.util.Objects;
import java.util.function.Function;
import java.util.stream.Collectors;

import interfaces.ExportBaseInterface;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;

import static services.ExcelUtility.ExportExcel;

public abstract class ExportCsv {

	public static final char FILE_DELIMITER = ';';
	public static final char QUOTE = '"';
	
	public static <T extends ExportBaseInterface> byte[] generateCsv(List<T> entitiesToWrite) {
		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(baos));

		T t = entitiesToWrite.get(0);
		List<String> header = Arrays.stream(t.getClass().getDeclaredFields()).sorted(orderComparator())
				 															 .map(getLabelFunction())
																			 .filter(Objects::nonNull)
																			 .collect(Collectors.toList());
		CSVFormat format = CSVFormat.DEFAULT.builder().setHeader(header.toArray(new String[0]))
													  .setDelimiter(FILE_DELIMITER)
													  .setQuote(QUOTE)
													  .build();

		try(CSVPrinter csvPrinter = new CSVPrinter(writer, format)){
			for(T e : entitiesToWrite) {
				List<Field> fields = Arrays.asList(e.getClass().getDeclaredFields());
				List<String> values = fields.stream().sorted(orderComparator())
													 .map(getValueFunction(e))
													 .filter(Objects::nonNull)
													 .collect(Collectors.toList());
				csvPrinter.printRecord(values);
			};
			
			csvPrinter.flush();
		} catch (Exception e) {
			throw new RuntimeException("Error in CSV creation",e);
		} 
		
		return baos.toByteArray();
	}

	/**
	 * @param <T>
	 * @param e
	 * @return
	 */
	private static <T extends ExportBaseInterface> Function<? super Field, ? extends String> getValueFunction(T e) {
		return f -> {
			try {
				f.setAccessible(true);
				ExportExcel annotation = f.getAnnotation(ExportExcel.class);
				if(annotation != null) {
					String value = "";
					Object valueObject = null;
					try {
						valueObject = f.get(e);
					} catch (IllegalArgumentException | IllegalAccessException e1) {
						//TODO LOGGER
					}
					if(valueObject != null)
						value = valueObject.toString();
					
					return value;
				}
			} finally {
				f.setAccessible(false);
			}
			return null;
		};
	}

	/**
	 * @return
	 */
	private static Function<? super Field, ? extends String> getLabelFunction() {
		return f -> {
			try {
				f.setAccessible(true);
				ExportExcel annotation = f.getAnnotation(ExportExcel.class);
				if(annotation != null)
					return annotation.label();
			} finally {
				f.setAccessible(false);
			}
			return null;
		};
	}

	/**
	 * @return
	 */
	private static Comparator<? super Field> orderComparator() {
		return (f1,f2) ->{
			int f1Order = -1;
			int f2Order = -1;
			
			ExportExcel annotation1 = f1.getAnnotation(ExportExcel.class);
			if(annotation1 != null) {
				f1Order = annotation1.order();				
			}
			ExportExcel annotation2 = f2.getAnnotation(ExportExcel.class);
			if(annotation2 != null) {
				f2Order = annotation2.order();				
			}
			
			return Integer.compare(f1Order, f2Order);
		};
	}

}
