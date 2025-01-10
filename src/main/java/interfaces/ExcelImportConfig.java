package interfaces;

import org.apache.poi.ss.usermodel.CellType;
public interface ExcelImportConfig {
	

	String getColumnTargetField();
	String getColumnTitle();
	CellType getColumnType();
	Boolean getRequired();
	Integer getColumnOrder();
	Integer getStartRow();
	Integer getStartColumn();
	Object getId();
}
