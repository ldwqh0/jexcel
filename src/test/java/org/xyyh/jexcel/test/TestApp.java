package org.xyyh.jexcel.test;

import java.util.ArrayList;
import java.util.List;

import org.xyyh.jexcel.core.ExcelMapper;
import org.xyyh.jexcel.test.entity.User;

public class TestApp {

	public static void main(String[] args) {
		
		// 新增一个 excel导出对象
		ExcelMapper mapper = new ExcelMapper();
		
		// 生成数据队列
		List<User> users = new ArrayList<User>();
		User user = new User();
		users.add(user);

		// 导出数据
		mapper.toExcel(users);
	}

}
