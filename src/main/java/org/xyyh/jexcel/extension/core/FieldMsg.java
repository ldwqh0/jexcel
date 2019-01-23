package org.xyyh.jexcel.extension.core;

import org.xyyh.jexcel.annotations.Col;
import org.xyyh.jexcel.annotations.ColIgnore;

import java.lang.reflect.Field;
import java.lang.reflect.Method;

public class FieldMsg {
    private Field field;
    private Method getMethod;
    private Method setMethod;
    private Col col;
    private ColIgnore colIgnore;
    private String colFrom;
    private String colIgnoreFrom;

    public Field getField() {
        return field;
    }

    public void setField(Field field) {
        this.field = field;
    }

    public Method getGetMethod() {
        return getMethod;
    }

    public void setGetMethod(Method getMethod) {
        this.getMethod = getMethod;
    }

    public Method getSetMethod() {
        return setMethod;
    }

    public void setSetMethod(Method setMethod) {
        this.setMethod = setMethod;
    }

    public Col getCol() {
        return col;
    }

    public void setCol(Col col) {
        this.col = col;
    }

    public ColIgnore getColIgnore() {
        return colIgnore;
    }

    public void setColIgnore(ColIgnore colIgnore) {
        this.colIgnore = colIgnore;
    }

    public String getColFrom() {
        return colFrom;
    }

    public void setColFrom(String coleFrom) {
        this.colFrom = coleFrom;
    }

    public String getColIgnoreFrom() {
        return colIgnoreFrom;
    }

    public void setColIgnoreFrom(String colIgnoreFrom) {
        this.colIgnoreFrom = colIgnoreFrom;
    }
}
