package org.xyyh.jexcel.core;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

public class ExcelMapper {

	public <T> Object toExcel(List<T> t) {
		// TODO retun a excel object
		return t;
	}

	public <T> void writeToExcel(List<T> t, OutputStream stream) {
		// TODO Write to excel;
	}

	public <T> List<T> parse(InputStream stream, List<T> datas, Class<T> cls) {
		// TODO 将数据写入到excel
		return datas;
	}
}
