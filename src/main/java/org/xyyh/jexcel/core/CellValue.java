package org.xyyh.jexcel.core;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;

public class CellValue<T> {

	private CellType cellType;

	private CellStyle cellStyle;

	private T value;

	public CellType getCellType() {
		return cellType;
	}

	public void setCellType(CellType cellType) {
		this.cellType = cellType;
	}

	public CellStyle getCellStyle() {
		return cellStyle;
	}

	public void setCellStyle(CellStyle cellStyle) {
		this.cellStyle = cellStyle;
	}

	public T getValue() {
		return value;
	}

	public void setValue(T value) {
		this.value = value;
	}

}
