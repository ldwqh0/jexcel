package org.xyyh.jexcel.core;

import java.util.List;

import org.apache.poi.ss.usermodel.CellValue;

public interface RowMapper<T> {

	/**
	 * 
	 * @param data
	 * @param index
	 * @return
	 */
	@Deprecated
	public <D> SimpleCellValue<?> getData(int rowNum, T data, int index);

	/**
	 * 获取表格的头信息
	 * 
	 * @return
	 */
	public List<String> getHeaders();

	@Deprecated
	public int getColumnCount(T data);

	public int getColumnCount();

	/**
	 * 根据列号从对象钟获取指定数据
	 * 
	 * @param colIndex
	 * @param data
	 * @return
	 */
	@Deprecated
	public SimpleCellValue<?> getCellValueD(int colIndex, T data);

	public CellValue getCellValue(int colIndex, T data);
}
