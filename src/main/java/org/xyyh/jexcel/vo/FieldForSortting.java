package org.xyyh.jexcel.vo;

import java.lang.reflect.Field;

/**
 * 属性排序
 */
public class FieldForSortting {
    private Field field;
    private String fieldName;
    private int index;

    public FieldForSortting() {
    }

    public FieldForSortting(Field field, String fieldName, int index) {
        this.field = field;
        this.fieldName = fieldName;
        this.index = index;
    }

    public Field getField() {
        return field;
    }

    public void setField(Field field) {
        this.field = field;
    }

    public String getFieldName() {
        return fieldName;
    }

    public void setFieldName(String fieldName) {
        this.fieldName = fieldName;
    }

    public int getIndex() {
        return index;
    }

    public void setIndex(int index) {
        this.index = index;
    }
}