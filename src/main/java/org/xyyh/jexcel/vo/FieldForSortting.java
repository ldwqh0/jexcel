package org.xyyh.jexcel.vo;


import java.lang.reflect.Field;
import java.util.Comparator;

/**
 * @author: max
 * @Date: 2018/12/17 0017 16:54
 * @Description: 所有属性排序
 */
public class FieldForSortting implements Comparator<FieldForSortting> {
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

    @Override
    public int compare(FieldForSortting o1, FieldForSortting o2) {
        return Integer.compare(o1.index, o2.index);
    }
}