package org.xyyh.jexcel.test.entity;

import org.xyyh.jexcel.annotations.Col;
import org.xyyh.jexcel.annotations.ColIgnore;
import org.xyyh.jexcel.annotations.Sheet;

@Sheet(name = "人员信息表")
public class User {

	@Col(name = "姓名")
	private String name;

	@Col(name = "年龄")
	private Long age;
	
	@ColIgnore
	private String password;

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public Long getAge() {
		return age;
	}

	public void setAge(Long age) {
		this.age = age;
	}

}
