package org.xyyh.jexcel.test;

import org.junit.Test;
import org.xyyh.jexcel.core.ExcelMapper;
import org.xyyh.jexcel.core.ObjectRowMapper;
import org.xyyh.jexcel.test.entity.Student;

import java.io.*;
import java.text.ParseException;
import java.util.*;

/**
 * @author: max
 * @Date: 2018/11/30 0030 11:45
 * @Description: 测试提交
 */
public class TestExcel {

	@Test
	public void Test() throws IOException {
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
		for (int i = 0; i < 3; i++) {
			Student student = new Student();
			student.setAge((i+1)*10L);
			student.setName("pangwa" + i + "");
			student.setScore(((double) (i + 5)));
			list.add(student);
		}
		String path = "d:/test1.xls";
        excelMapper.toExcel(list,path);
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
		List<Map> list = new ArrayList<>();
		for (int i = 0; i < 3; i++) {
			Map<String, Object> map = new LinkedHashMap<>();
			map.put("name", "" + "mm" + i);
			map.put("age", 12 + i);
			map.put("score", i);
			list.add(map);
		}
		String path = "d:/test.xls";
		//保存到指定路径
//		excelMapper.toExcel(list, path);
		//保存到指定文件
//		excelMapper.toExcel(list, new File("d:/map.xls"));
		//保存到指定输出流
		FileOutputStream fileOutputStream = new FileOutputStream("d:/map.xls");
		excelMapper.toExcel(list,fileOutputStream);

	}

	@Test
	public void testUserDefine() {
		Student stu = new Student();
		stu.setScore(1d);
		new ObjectRowMapper<>(Student.class);
	}

	@Test
	public void testImport() throws IOException, IllegalAccessException, ParseException, InstantiationException {
		ExcelMapper ex = new ExcelMapper();
		FileInputStream fis = new FileInputStream("d:/map.xls");
		//导入生成实体
//        List<Student> list = new ArrayList<>();
//        list = ex.parse(fis, list, Student.class);
		//导入生成map
		List<Map> list1 = new ArrayList<>();
		List<Map> mapList = ex.parse(fis, list1, Map.class);
		System.out.println(mapList);
//		System.out.println(list.get(0));
	}
}
