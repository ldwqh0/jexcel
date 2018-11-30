package org.xyyh.jexcel.core;

public interface RowMapper<T> {

	/**
	 * 
	 * @param data
	 * @param index
	 * @return
	 */
	public <D> CellValue getData(T data, int index);

	public int getColumnCount(T data);
}
