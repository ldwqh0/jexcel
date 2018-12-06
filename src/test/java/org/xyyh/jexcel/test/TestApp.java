package org.xyyh.jexcel.test;

import java.util.ArrayList;
import java.util.List;

import org.xyyh.jexcel.core.ExcelMapper;
import org.xyyh.jexcel.test.entity.User;

public class TestApp {

	public static void main(String[] args) {
		
		// ����һ�� excel��������
		ExcelMapper mapper = new ExcelMapper();
		
		// �������ݶ���
		List<User> users = new ArrayList<User>();
		User user = new User();
		users.add(user);

		// ��������
		mapper.toWorkBook(users);
	}

}
