package org.xyyh.jexcel.core;

import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellValue;

/**
 * 固定列标头的rowMapper
 * 
 * @author LiDong
 *
 */
public class FixedColumnMapRowMapper implements RowMapper<Map<String, Object>> {

	private List<String> headers;
	private List<String> keys;

	public FixedColumnMapRowMapper(List<String> headers, List<String> keys) {
		this.headers = headers;
		this.keys = keys;
	}

	protected FixedColumnMapRowMapper() {

	}

	public void setHeaders(List<String> headers) {
		this.headers = headers;
	}

	public void setKeys(List<String> keys) {
		this.keys = keys;
	}

	@Override
	public List<String> getHeaders() {
		return this.headers;
	}

	@Override
	public int getColumnCount() {
		return this.headers.size();
	}

	@Override
	public CellValue getCellValue(int colIndex, Map<String, Object> data) {
		String key = keys.get(colIndex);
		Object value = data.get(key);
		if (value instanceof Number) {
			return new CellValue(((Number) value).doubleValue());
		} else if (value instanceof Boolean) {
			return Boolean.TRUE.equals(value) ? CellValue.TRUE : CellValue.FALSE;
		} else {
			return new CellValue(String.valueOf(value));
		}
	}

}
