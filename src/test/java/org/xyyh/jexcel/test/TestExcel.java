package org.xyyh.jexcel.test;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.Test;
import org.xyyh.jexcel.core.ExcelMapper;

/**
 * @author: max
 * @Date: 2018/11/30 0030 11:45
 * @Description: 测试提交
 */
public class TestExcel {
	public static void main(String[] args) {
		System.out.println("Hello World");
	}

	@Test
	public void testMapExport() throws FileNotFoundException, IOException {
		List<Map<String, Object>> datas = new ArrayList<>();
		for (int i = 0; i < 10; i++) {
			Map<String, Object> data = new HashMap<>();
			data.put("index", i);
			data.put("name", "name" + i);
			data.put("age", i + 10);
			datas.add(data);
		}
		ExcelMapper mapper = new ExcelMapper();
		mapper.toExcel(datas, "d:/text.xlsx");
	}
}
