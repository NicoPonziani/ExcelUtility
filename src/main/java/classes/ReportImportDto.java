package classes;

import enums.StatusRowEnum;

import java.io.Serial;
import java.io.Serializable;

public class ReportImportDto implements Serializable{
	
	@Serial
	private static final long serialVersionUID = 1L;

	private String idRow;
	private String message; 
	private int rowIndex;
	private Integer columnIndex;
	private StatusRowEnum status;
	
	public ReportImportDto() {super();}
	
	public ReportImportDto(String errorMessage) {
		this.message = errorMessage;
	}
	
	public ReportImportDto(String errorMessage, int rowIndex, Integer columnIndex) {
		this(errorMessage);
		this.rowIndex = rowIndex;
		this.columnIndex = columnIndex;
		this.status = StatusRowEnum.ERROR;
	}
	
	public ReportImportDto(String errorMessage, int rowIndex, Integer columnIndex, StatusRowEnum status) {
		this(errorMessage,rowIndex,columnIndex);
		this.status = status;
	}
	
	public ReportImportDto(String errorMessage, int rowIndex, Integer columnIndex, StatusRowEnum status, String idRow) {
		this(errorMessage,rowIndex,columnIndex, status);
		this.idRow = idRow;
	}

	public String getMessage() {
		return message;
	}

	public void setMessage(String errorMessage) {
		this.message = errorMessage;
	}

	public int getRowIndex() {
		return rowIndex;
	}

	public void setRowIndex(int indexColumn) {
		this.rowIndex = indexColumn;
	}

	public Integer getColumnIndex() {
		return columnIndex;
	}

	public void setColumnIndex(Integer columnIndex) {
		this.columnIndex = columnIndex;
	}

	public StatusRowEnum getStatus() {
		return status;
	}

	public void setStatus(StatusRowEnum status) {
		this.status = status;
	}

	public String getIdRow() {
		return idRow;
	}

	public void setIdRow(String idRow) {
		this.idRow = idRow;
	}
}
