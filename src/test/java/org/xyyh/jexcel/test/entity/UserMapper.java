package org.xyyh.jexcel.test.entity;

import java.util.List;

import org.xyyh.jexcel.core.SimpleCellValue;
import org.apache.poi.ss.usermodel.CellValue;
import org.xyyh.jexcel.core.RowMapper;

public class UserMapper implements RowMapper<User> {

	public <D> SimpleCellValue getData(User data, int index) {
		if (index == 1) {
			SimpleCellValue cellValue = new SimpleCellValue<>();
			cellValue.setCellStyle(null);
			cellValue.setCellType(null);
			cellValue.setValue(data.getName());
			return cellValue;
		} else {
			return null;
		}
	}

	@Override
	public int getColumnCount(User data) {
		return 3;
	}

	@Override
	public <D> SimpleCellValue getData(int rowNum, User data, int index) {
		// TODO Auto-generated method stub
		return null;
	}

	public <D> SimpleCellValue<?> getData(int colIndex, User data) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public List<String> getHeaders() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public int getColumnCount() {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public SimpleCellValue<?> getCellValueD(int colIndex, User data) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public CellValue getCellValue(int colIndex, User data) {
		// TODO Auto-generated method stub
		return null;
	}

}
