package org.xyyh.jexcel.entity;

import java.util.Date;

/**
 * @author: max
 * @Date: 2018/12/3 0003 16:42
 * @Description: 学生对象
 */
public class Student {

    private Integer id;

    private String name;

    private Double score;

    private Date birthday;

    public Integer getId() {
        return id;
    }

    public void setId(Integer id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Double getScore() {
        return score;
    }

    public void setScore(Double score) {
        this.score = score;
    }

    public Date getBirthday() {
        return birthday;
    }

    public void setBirthday(Date birthday) {
        this.birthday = birthday;
    }
}
