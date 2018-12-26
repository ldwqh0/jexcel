package org.xyyh.jexcel.test.entity;


import org.xyyh.jexcel.annotations.Col;
import org.xyyh.jexcel.annotations.Sheet;

/**
 * @author: max
 * @Date: 2018/12/3 0003 16:42
 * @Description: 学生对象
 */
@Sheet(name = "学生表")
public class Student extends User {

	@Col(name = "成绩",sort = 2)
	private Double score;

	public Double getScore() {
		return score;
	}

	public void setScore(Double score) {
		this.score = score;
	}
}
