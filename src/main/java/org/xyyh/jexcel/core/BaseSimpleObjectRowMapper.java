package org.xyyh.jexcel.core;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;

import java.lang.reflect.Field;
import java.util.List;

public class BaseSimpleObjectRowMapper implements RowMapper<Object> {

	@Override
	public <D> SimpleCellValue getData(int rowNum, Object data, int index) {
		// TODO Auto-generated method stub
		SimpleCellValue<Object> cellValue = new SimpleCellValue<>();
		cellValue.setCellStyle(null);
		cellValue.setCellType(CellType.STRING);
		Field[] fields = data.getClass().getDeclaredFields();
		String name = fields[index].getName();
		Field field;
		try {
			field = data.getClass().getDeclaredField(name);
			field.setAccessible(true);
			Object value = field.get(data);
			if (value != null) {
				if (rowNum > 0) {
					cellValue.setValue(value);
				} else {
					cellValue.setValue(name);
				}
			} else {
				cellValue.setValue("");
			}
		} catch (NoSuchFieldException | IllegalAccessException e) {
			e.printStackTrace();
		}
		return cellValue;
	}

	@Override
	public int getColumnCount(Object data) {
		// TODO Auto-generated method stub
		Field[] fields = data.getClass().getDeclaredFields();
		int length = fields.length;
		return length;
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

	public <D> SimpleCellValue<?> getData(int colIndex, Object data) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public SimpleCellValue<?> getCellValueD(int colIndex, Object data) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public CellValue getCellValue(int colIndex, Object data) {
		// TODO Auto-generated method stub
		return null;
	}

}
