package org.xyyh.jexcel.extension.core;

import org.apache.poi.ss.usermodel.CellValue;
import org.xyyh.jexcel.annotations.Col;
import org.xyyh.jexcel.annotations.ColIgnore;
import org.xyyh.jexcel.core.RowMapper;
import org.xyyh.jexcel.extension.util.MyClassUtil;

import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class ObjectRowMapper<T>  implements RowMapper<T> {

    private Class<T> tClass;

    private List<String> headers;

    private List<FieldMsg> orderList;

    private List<Field> keys;

    private int columnCount;

    public  ObjectRowMapper(Class<T> tClass){
         this.tClass =tClass;
         this.headers =new ArrayList<String>();
         this.keys =new ArrayList<Field>();
         //按照注解排好序后的字段信息，可以注解在get set 方法上面
         this.orderList = MyClassUtil.getAllFields(tClass);
    }
    /*
    *   map.put("field",field);
        map.put("getMethod",getMethod);
        map.put("setMethod",setMethod);
        map.put("colIgnoreFrom",colIgnoreFrom);
        map.put("colFrom",colFrom);
        map.put("col",col);
        map.put("colIgnore",colIgnore);
    * */
    @Override
    public List<String> getHeaders() {
        if (!headers.isEmpty()){
            return  headers;
        }
        for (int i = 0; i < orderList.size(); i++) {
            FieldMsg fieldMsg =orderList.get(i);
           String header ="";
            ColIgnore colIgnore = fieldMsg.getColIgnore();
            //实体倒表映射，只取get方法的忽略 和字段的忽略
            if (colIgnore!=null&&colIgnore.value()
                    &&!"setMethod".equals(fieldMsg.getColIgnoreFrom())){
                continue;
            }
            Col col = fieldMsg.getCol();
            //实体倒表映射，只取get方法的 和字段的
            if (col!=null&&!"setMethod".equals(fieldMsg.getColFrom())){
                String colName =col.name();
                header ="".equals(colName)?fieldMsg.getField().getName():colName;
                headers.add(header);
                keys.add(fieldMsg.getField());
                continue;
            }
            //get方法的 和字段有任何注解则,并且存在get 方法，用字段名做表头
            if (fieldMsg.getGetMethod()!=null){
                header =fieldMsg.getField().getName();
                headers.add(header);
                keys.add(fieldMsg.getField());
            }
        }
        return headers;
    }

    @Override
    public int getColumnCount() {
        if (headers.isEmpty()){
            getHeaders();
        }
        return headers.size();
    }

    @Override
    public CellValue getCellValue(int colIndex, Object data){
        if (keys.isEmpty()){
            getHeaders();
        }
        Field field =  keys.get(colIndex);
        field.setAccessible(true);

        Object object = null;
        try {
            object = field.get(data);
        } catch (IllegalAccessException e) {
            System.out.println("设置第"+colIndex+"列失败，初始化化为空字符串");
            return new CellValue("");
        }
        field.setAccessible(false);
        CellValue  cellValue =null;
        //数字类型能否转换为数字
        if (Number.class.isInstance(object)){
            cellValue =new CellValue(((Number) object).doubleValue());
        }else if (String.class.isInstance(object)){
            cellValue =new CellValue((String) object);
        }else if (Date.class.isInstance(object)){
            SimpleDateFormat simpleDateFormat =new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            cellValue =new CellValue(simpleDateFormat.format(object));
        }else if (LocalDate.class.isInstance(object)){
            DateTimeFormatter fmt = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
            LocalDate localDate =(LocalDate)object;
            cellValue =new CellValue(localDate.format(fmt));
        }else if (ZonedDateTime.class.isInstance(object)){
            ZonedDateTime zdt=(ZonedDateTime)object;
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
            cellValue =new CellValue(zdt.format(formatter));
        }else {
            cellValue =new CellValue(String.valueOf(object));
        }
        return cellValue;
    }

    public Class<T> gettClass() {
        return tClass;
    }

    public List<FieldMsg> getOrderList() {
        return orderList;
    }

    public List<Field> getKeys() {
        return keys;
    }
}
