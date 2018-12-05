package org.xyyh.jexcel.core;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;

public class SimpleCellValue<T> {

	private CellType cellType;

	private CellStyle cellStyle;

	private T value;

	public SimpleCellValue() {
		// 默认无参构造器
	}

	// 将任意对象转换为一个 CellValue
	@SuppressWarnings("unchecked")
	public SimpleCellValue(Object value2) {
		this.value = (T) value2;
		this.cellType = CellType.NUMERIC;
		this.cellType = CellType.STRING;
	}

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
