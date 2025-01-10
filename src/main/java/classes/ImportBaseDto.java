package classes;

import org.apache.commons.lang3.StringUtils;
import services.ExcelUtility;

public class ImportBaseDto {
	
	private int nrRow;
	
	public ImportBaseDto() {
		super();
	}

	public int getNrRow() {
		return nrRow;
	}

	public void setNrRow(int nrRow) {
		this.nrRow = nrRow;
	}
	
	public Integer getAnnotationOrder(String field) {
		if(StringUtils.isBlank(field))return null;
		ExcelUtility.ImportExcel declaredAnnotation = null;
		try {
			declaredAnnotation = this.getClass().getDeclaredField(field).getDeclaredAnnotation(ExcelUtility.ImportExcel.class);
		} catch (NoSuchFieldException | SecurityException e) {
			return null;
		}
		return declaredAnnotation.order();
	}
}
