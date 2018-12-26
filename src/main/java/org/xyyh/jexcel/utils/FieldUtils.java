package org.xyyh.jexcel.utils;

import org.xyyh.jexcel.annotations.Col;
import org.xyyh.jexcel.annotations.ColIgnore;
import org.xyyh.jexcel.vo.FieldForSortting;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;

/**
 * @author: max
 * @Date: 2018/12/17 0017 16:54
 * @Description: 用于获取所有属性
 */
public class FieldUtils {


    /**
     * 给实体类属性值制定顺序
     * @param clazz
     * @return
     */
    public static List<FieldForSortting> sortFieldByAnno(Class<?> clazz)  {
        List<FieldForSortting> list = new ArrayList<>();
        FieldUtils.getAllFields(list,clazz);
        list.sort(Comparator.comparingInt(FieldForSortting::getIndex));
        return list;
    }

    /**
     * 递归调用，获取本类所有属性
     * @param list
     * @param clazz
     */
    private static void getAllFields(List<FieldForSortting> list, Class<?> clazz){
        if(clazz != Object.class){
            returnClassField(list, clazz);
            Class supperClazz = clazz.getSuperclass();
            getAllFields(list,supperClazz);
        }
    }

    /**
     * 获取当前类所有属性
     * @param list
     * @param clazz
     */
    private static void returnClassField(List<FieldForSortting> list, Class<?> clazz) {
        Field[] declaredFields = clazz.getDeclaredFields();
        for (Field declaredField : declaredFields) {
            FieldForSortting ffs = new FieldForSortting();
            declaredField.setAccessible(true);
            Annotation[] annotations = declaredField.getAnnotations();
            for (Annotation annotation : annotations) {
                if (annotation instanceof ColIgnore){
                    break;
                } else if(annotation instanceof Col){
                    Col column = (Col)annotation;
                    String name = column.name();
                    int sort = column.sort();
                    ffs.setFieldName("".equals(name)?declaredField.getName():name);
                    ffs.setIndex(sort >= 0?sort:-1);
                    ffs.setField(declaredField);
                }
                list.add(ffs);
            }
        }
    }
}
