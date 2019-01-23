package org.xyyh.jexcel.extension.util;

import org.xyyh.jexcel.annotations.Col;
import org.xyyh.jexcel.annotations.ColIgnore;
import org.xyyh.jexcel.extension.core.FieldMsg;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.*;

public class MyClassUtil {





    /**
     * 获取包括父类的所有字段,和对应字段的get 和set 方法
     * @param clazz 要获取字段的类
     * @return  list<Map<key,obeject>> 对应的分别是：
     *           字段名-field，get字段名 -getMethod set字段名 -setMethod
     * */
    public static List<FieldMsg> getAllFields(Class clazz){
        Class superClass =clazz.getSuperclass();
        List<FieldMsg> list =new ArrayList<>();
        boolean flag =superClass!=null;

        //有父类的时候进入，获取父类的字段和方法
        while (flag){
            Field[] fields =superClass.getDeclaredFields();
            for (int i = 0; i <fields.length ; i++) {
                Map<String,Object> map =new HashMap<>();
                list.add(getFieldMsg(fields[i],superClass));
            }
            superClass =superClass.getSuperclass();
            flag =superClass!=null;
        }
        Field[] fields = clazz.getDeclaredFields();
        for (int i = 0; i <fields.length ; i++) {
           list.add(getFieldMsg(fields[i],clazz)) ;
        }
        return sortMap(list);
    }

    /**
     * 获取一个字段的信息，包括字段的get set方法
     * */
    public static FieldMsg getFieldMsg(Field field, Class clazz){
         if (field==null){
                return  null;
            }
        FieldMsg fieldMsg =new FieldMsg();
        String fieldName =field.getName();
        String getMethodName ="get"+fieldName.substring(0,1).toUpperCase()+fieldName.substring(1);
        String setMethodName ="set"+fieldName.substring(0,1).toUpperCase()+fieldName.substring(1);
        Method getMethod=null;
        Method setMethod=null;
        try {
            getMethod = clazz.getDeclaredMethod(getMethodName,null);
            setMethod =clazz.getDeclaredMethod(setMethodName,field.getType());
        } catch (NoSuchMethodException e) {
            System.out.println(fieldName+"没有get 或 set 方法，已跳过方法获取");
        }
        Col col =field.getAnnotation(Col.class);
        String colFrom ="field";
        if (col==null&&getMethod!=null){
            col =getMethod.getAnnotation(Col.class);
            colFrom ="getMethod";
        }
        if (col==null&&setMethod!=null){
            col =setMethod.getAnnotation(Col.class);
            colFrom ="setMethod";
        }

        ColIgnore colIgnore =field.getAnnotation(ColIgnore.class);
        String colIgnoreFrom ="field";
        if (colIgnore==null&&getMethod!=null){
            colIgnore =getMethod.getAnnotation(ColIgnore.class);
            colIgnoreFrom ="getMethod";
        }
        if (colIgnore==null&&setMethod!=null){
            colIgnore =setMethod.getAnnotation(ColIgnore.class);
            colIgnoreFrom ="setMethod";
        }
        fieldMsg.setField(field);
        fieldMsg.setGetMethod(getMethod);
        fieldMsg.setSetMethod(setMethod);
        fieldMsg.setColIgnoreFrom(colIgnoreFrom);
        fieldMsg.setColFrom(colFrom);
        fieldMsg.setCol(col);
        fieldMsg.setColIgnore(colIgnore);
       return  fieldMsg;
    }

    /**
     * 根据注解信息排序,顺便吧注解放入map
     * */
    public static List<FieldMsg> sortMap(List<FieldMsg> list){
        list.sort(new Comparator<FieldMsg>() {
            @Override
            public int compare(FieldMsg map1, FieldMsg map2) {
                int m =1;
                 Col col1 =(Col)map1.getCol();
                 if (col1!=null){
                     Col col2 =(Col)map2.getCol();
                     if (col2!=null){
                         return  col1.sort()-col2.sort();
                     }
                 }
                return m;
            }
        });
        return list;
    }



}
