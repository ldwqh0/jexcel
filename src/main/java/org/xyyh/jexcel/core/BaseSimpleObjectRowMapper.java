package org.xyyh.jexcel.core;

import java.util.List;

import org.apache.poi.ss.usermodel.CellValue;

public class BaseSimpleObjectRowMapper<T> implements RowMapper<T> {

//	@Override
//	public <D> SimpleCellValue<?> getData(int rowNum, T data, int index) {
//		SimpleCellValue<Object> cellValue = new SimpleCellValue<>();
//		cellValue.setCellStyle(null);
//		cellValue.setCellType(CellType.STRING);
//		Field[] fields = data.getClass().getDeclaredFields();
//		String name = fields[index].getName();
//		Field field;
//		try {
//			field = data.getClass().getDeclaredField(name);
//			field.setAccessible(true);
//			Object value = field.get(data);
//			if (value != null) {
//				if (rowNum > 0) {
//					cellValue.setValue(value);
//				} else {
//					cellValue.setValue(name);
//				}
//			} else {
//				cellValue.setValue("");
//			}
//		} catch (NoSuchFieldException | IllegalAccessException e) {
//			e.printStackTrace();
//		}
//		return cellValue;
//	}

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
	public CellValue getCellValue(int colIndex, T data) {
		// TODO Auto-generated method stub
		return null;
	}
}
