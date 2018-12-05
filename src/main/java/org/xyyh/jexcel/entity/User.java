package org.xyyh.jexcel.entity;

import org.xyyh.jexcel.annotations.Col;
import org.xyyh.jexcel.annotations.ColIgnore;
import org.xyyh.jexcel.annotations.Sheet;

@Sheet(name = "��Ա��Ϣ��")
public class User {

	@Col(name = "����")
	private String name;

	@Col(name = "����")
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
