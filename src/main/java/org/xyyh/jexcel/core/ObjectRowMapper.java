package org.xyyh.jexcel.core;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.CellValue;
import org.xyyh.jexcel.utils.FieldUtils;
import org.xyyh.jexcel.vo.FieldForSortting;

/**
 * 转换对象的{@link RowMapper}
 * 
 * @author LiDong
 *
 * @param <T>
 */
public class ObjectRowMapper<T> implements RowMapper<T> {

	private Class<T> tClass;

	private List<String> headers = new ArrayList<String>();

	private List<Field> keys = new ArrayList<>();

	private int columnCount;

	public ObjectRowMapper(Class<T> tClass) {
		try {
			this.tClass = tClass;
			
			// 获取所有get/set方法
			// 获取所有字段
			// 获取所有组件
//			Method[] methods = tClass.getMethods();
//			for (int i = 0; i < methods.length; i++) {
//				System.out.println(methods[i]);
//			}
			List<FieldForSortting> ffs = FieldUtils.sortFieldByAnno(tClass);
			List<String> fieldNames = new ArrayList<>();
			List<Field> fields = new ArrayList<>();
			for (FieldForSortting ff : ffs) {
				String fieldName = ff.getFieldName();
				Field field = ff.getField();
				fieldNames.add(fieldName);
				fields.add(field);
			}
			setHeaders(fieldNames);
			setKeys(fields);
		} catch (Exception e) {
			// TODO: handle exception
		}
	}

	public void setHeaders(List<String> headers) {
		this.headers = headers;
	}

	public void setKeys(List<Field> keys) {
		this.keys = keys;
	}

	@Override
	public List<String> getHeaders() {
		return headers;
	}

	@Override
	public int getColumnCount() {
		return headers.size();
	}

	@Override
	public CellValue getCellValue(int colIndex, T data) {
		Field field = keys.get(colIndex);
		field.setAccessible(true);
		try {
			Object value = field.get(data);
			if (value instanceof Number) {
				return new CellValue(((Number) value).doubleValue());
			} else if (value instanceof Boolean) {
				return Boolean.TRUE.equals(value) ? CellValue.TRUE : CellValue.FALSE;
			} else {
				if(value != null){
					return new CellValue(String.valueOf(value));
				}else{
					return new CellValue(String.valueOf(""));
				}
			}
		} catch (IllegalAccessException e) {
			e.printStackTrace();
		}
		return new CellValue(String.valueOf(""));
	}

}
