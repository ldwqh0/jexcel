package org.xyyh.jexcel.test;

import org.junit.Test;
import org.xyyh.jexcel.core.ExcelMapper;
import org.xyyh.jexcel.core.ObjectRowMapper;
import org.xyyh.jexcel.test.entity.Student;

import java.io.*;
import java.util.*;

/**
 * @author: max
 * @Date: 2018/11/30 0030 11:45
 * @Description: 测试提交
 */
public class TestExcel {
	@Test
	public void Test() {
		ExcelMapper excelMapper = new ExcelMapper();
//        List<Map> list = new ArrayList<>();
//        for (int i = 0; i < 3; i++) {
//            Map<Object,Object> map = new HashMap<>();
//            map.put("name",""+"mm"+i);
//            map.put("age",12+i);
//            map.put(12,++i);
//            list.add(map);
//        }
		OutputStream os = System.out;
//        excelMapper.WriteToExcel(list,os);

		List<Student> list = new ArrayList<>();
		;
		for (int i = 0; i < 3; i++) {
			Student student = new Student();
//			student.setId(i);
//			student.setBirthday(new Date());
			student.setName("pangwa" + i + "");
			student.setScore(((double) (i + 5)));
			list.add(student);
		}
//        excelMapper.writeToExcel(list,os);
		System.out.println("Hello World");
	}

	/**
	 * 输入map格式导出到制定路径
	 * 
	 * @throws IOException
	 */
	@Test
	public void testToExcelByMap() throws IOException {
		ExcelMapper excelMapper = new ExcelMapper();
		List<LinkedHashMap> list = new ArrayList<>();
		for (int i = 0; i < 3; i++) {
			Map<Object, Object> map = new LinkedHashMap<>();
			map.put("name", "" + "mm" + i);
			map.put("age", 12 + i);
			map.put("score", i);
			list.add((LinkedHashMap) map);
		}
		String path = "d:/test.xls";
		excelMapper.toExcel(list, path);

	}

	@Test
	public void testUserDefine() {
		Student stu = new Student();
		stu.setScore(1d);
		new ObjectRowMapper<>(Student.class);
	}

	@Test
	public void testImport() throws IOException {
		ExcelMapper ex = new ExcelMapper();
		FileInputStream fis = new FileInputStream("d:/test1.xls");
		//导入生成实体
        List<Student> list = new ArrayList<>();
        list = ex.parse(fis, list, Student.class);
		//导入生成map
//		List<Map> list1 = new ArrayList<>();
//		List<Map> mapList = ex.parse(fis, list1, Map.class);
//		System.out.println(mapList);
		System.out.println(list.size());
	}

}
