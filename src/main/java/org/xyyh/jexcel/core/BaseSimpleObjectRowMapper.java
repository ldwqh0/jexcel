package org.xyyh.jexcel.core;

import org.apache.poi.ss.usermodel.CellType;

import java.lang.reflect.Field;

public class BaseSimpleObjectRowMapper<T> implements RowMapper<T> {

	@Override
	public <D> CellValue getData(int rowNum, T data, int index) {
			CellValue<Object> cellValue = new CellValue<>();
			cellValue.setCellStyle(null);
			cellValue.setCellType(CellType.STRING);
			Field[] fields = data.getClass().getDeclaredFields();
			String name = fields[index].getName();
			Field field;
			try {
				field = data.getClass().getDeclaredField(name);
				field.setAccessible(true);
				Object value = field.get(data);
				if(value != null){
					if(rowNum >0){
						cellValue.setValue(value);
					}else{
						cellValue.setValue(name);
					}
				}else {
					cellValue.setValue("");
				}
			} catch (NoSuchFieldException | IllegalAccessException e) {
				e.printStackTrace();
			}
			return cellValue;
		}

	@Override
	public int getColumnCount(T data) {
		Field[] fields = data.getClass().getDeclaredFields();
		int length = fields.length;
		return length;
	}
}
