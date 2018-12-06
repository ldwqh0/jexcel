package org.xyyh.jexcel.core;

import java.util.List;

import org.apache.poi.ss.usermodel.CellValue;

public interface RowMapper<T> {



	/**
	 * 获取表格的头信息
	 * 
	 * @return
	 */
	public List<String> getHeaders();


	public int getColumnCount();

	
	public CellValue getCellValue(int colIndex, T data);
}
