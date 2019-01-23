package org.xyyh.jexcel.core;

import org.apache.poi.ss.usermodel.CellValue;

import java.util.List;

public interface RowMapper<T> {



	/**
	 * 获取表格的头信息
	 * 
	 * @return
	 */
	public List<String> getHeaders();


	public int getColumnCount();

	
	public CellValue getCellValue(int colIndex, T data) ;
}
