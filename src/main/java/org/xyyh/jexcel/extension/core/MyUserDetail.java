package org.xyyh.jexcel.extension.core;

import org.xyyh.jexcel.annotations.Col;

import java.io.Serializable;
import java.util.Date;

/**
 * 自定义客户信息
 * */
public class MyUserDetail  implements Serializable {


    private static final long serialVersionUID = -6186893015772300645L;
    @Col(name = "id")
    private String clientId;
    @Col(name = "秘钥")
    private String clientSecret;
    @Col(name = "姓名")
    private String name;
    @Col(name = "时间")
    private Date date;
    public String getClientId() {
        return clientId;
    }

    public void setClientId(String clientId) {
        this.clientId = clientId;
    }

    public String getClientSecret() {
        return clientSecret;
    }

    public void setClientSecret(String clientSecret) {
        this.clientSecret = clientSecret;
    }
    @Col(name = "名字")
    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }
}
