package org.xyyh.jexcel.core;

import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.CellValue;

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

	private List<String> keys = new ArrayList<>();

	private int columnCount;

	public ObjectRowMapper(Class<T> tClass) {
		try {
			this.tClass = tClass;
			
			// 获取所有get/set方法
			// 获取所有字段
			// 获取所有组件

			Method[] methods = tClass.getMethods();
			for (int i = 0; i < methods.length; i++) {
				System.out.println(methods[i]);
			}

		} catch (Exception e) {
			// TODO: handle exception
		}
	}

	@Override
	public List<String> getHeaders() {
		return headers;
	}

	@Override
	public int getColumnCount() {
		// TODO Auto-generated method stub
		return 0;
	}

	@Override
	public CellValue getCellValue(int colIndex, T data) {
//	
		return null;
	}

}
