package org.xyyh.jexcel.test.entity;

import java.util.List;

import org.xyyh.jexcel.core.RowMapper;
import org.xyyh.jexcel.core.SimpleCellValue;

public class UserMapper implements RowMapper<User> {

//	@Override
	public <D> SimpleCellValue<?> getData(User data, int index) {
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
	public org.apache.poi.ss.usermodel.CellValue getCellValue(int colIndex, User data) {
		// TODO Auto-generated method stub
		return null;
	}

}
