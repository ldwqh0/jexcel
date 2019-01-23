package org.xyyh.jexcel.extension.core;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.xyyh.jexcel.annotations.Col;

import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class SheetRowMapper {

    private  String  sheetName;

    private List<String> headers;
    //注意所有的时间格式字符串我转换成的localData
    private List<LinkedHashMap<String,Object>> dataList;

    private int columnCount;

    public SheetRowMapper(Sheet sheet){
        init(sheet);
    }


    public String getSheetName() {
        return sheetName;
    }

    public List<String> getHeaders() {
        return headers;
    }

    public List<LinkedHashMap<String,Object>> getDataList() {
        return dataList;
    }

    public int getColumnCount() {
        return columnCount;
    }

    public void reLoadBySheet(Sheet sheet){
        init(sheet);
    }



    public <T> List<T> getBeanist(Class<T> tClass){
        if (tClass.isAssignableFrom(Map.class)){
            return  (List<T>) dataList;
        }
        List<T> tList =new ArrayList<>();
        List<Map<String,Object>> list =getInitField(tClass);
        try {
            for (int i = 0; i <dataList.size() ; i++) {
                LinkedHashMap<String,Object> map1 =dataList.get(i);
                T t= tClass.newInstance();
                for (int j = 0; j < list.size(); j++) {
                    Map<String,Object> fieldMap=list.get(j);
                    String key =(String)fieldMap.get("header");
                    Field field =(Field)fieldMap.get("field");
                    Object value =map1.get(key);
                    field.setAccessible(true);
                    String fieldType =field.getType().getSimpleName();
                    //因为时间字符串我全转的localdate 说以是时间需要处理下
                    if (fieldType.equalsIgnoreCase("date")){
                        LocalDate localDate =(LocalDate)value;
                        System.out.println(localDate==null);
                        ZonedDateTime zdt = localDate.atStartOfDay(ZoneId.systemDefault());
                        Date date =Date.from(zdt.toInstant());
                        field.set(t,date);
                    }else if (fieldType.equalsIgnoreCase("ZonedDateTime")){
                        LocalDate localDate =(LocalDate)value;
                        ZonedDateTime zdt = localDate.atStartOfDay(ZoneId.systemDefault());
                        field.set(t,zdt);
                    }else {
                        field.set(t,value);
                    }
                    field.setAccessible(false);
                }
                tList.add(t);
            }
        }catch (Exception e){
            e.printStackTrace();
        }
        return tList;
    }


    /**
     * 根据注解信息排序,顺便吧注解放入map
     * @return map.put("header ", lisr < string >);
     *         map.put("data",List<LinkedHashMap<String,Object>>);
     *         map.put("name",sheetName);
     * */
    private void init(Sheet sheet){
        if (sheet==null){
            System.out.println("sheet为空");
            return;
        }
        headers =new ArrayList<>();
        dataList =new ArrayList<>();
        sheetName =sheet.getSheetName();
        for (int i = 0; i < sheet.getLastRowNum()+1; i++) {
            Row row = sheet.getRow(i);
            LinkedHashMap<String,Object> linkedHashMap =new LinkedHashMap<>();
            if (i==0){
                for (int j = 0; j <row.getPhysicalNumberOfCells(); j++) {
                    String value =String.valueOf(getCellValue(row.getCell(j)));
                    headers.add(value);
                }
            }else{
                for (int j = 0; j <row.getPhysicalNumberOfCells(); j++) {
                    linkedHashMap.put(headers.get(j),getCellValue(row.getCell(j)));
                }
                dataList.add(linkedHashMap);
            }
        }
        columnCount=headers.size();
    }

    private  Object getCellValue(Cell cell) {
        Object ret =null;
        CellType type = cell.getCellType();
        switch (type) {
            case BLANK:
                ret = "";
                break;
            case BOOLEAN:
                ret = cell.getBooleanCellValue();
                break;
            case STRING:
                String timeStr = cell.getRichStringCellValue().getString().trim();
                if (timeStr.matches("\\d{4}-\\d{2}-\\d{2}\\s+\\d{2}:\\d{2}:\\d{2}")){
                    String patter = timeStr.matches("\\d{4}-\\d{2}-\\d{2}\\s+\\d{2}:\\d{2}:\\d{2}")?
                            "yyyy-MM-dd HH:mm:ss":"yyyy-MM-dd";
                    DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern(patter);
                    ret = LocalDate.parse(timeStr,dateTimeFormatter);

                }else {
                    ret=timeStr;
                }
                break;
            case NUMERIC:
                ret =cell.getNumericCellValue();
                break;
            case FORMULA://公式不处理
                break;
            case ERROR://错误不处理
                ret = null;
                break;
            default:
                ret =null;
        }
        return ret;
    }


    /**
     *list
     * */
    private   <T> List<Map<String,Object>> getInitField(Class<T> tClass){
        ObjectRowMapper<T> objectRowMapper =new ObjectRowMapper<>(tClass);
         //tClass字段信息
        List<FieldMsg> list = objectRowMapper.getOrderList();
         //需要初始的字段
        List<Map<String,Object>> list3 =new ArrayList<>();

        for (int i = 0; i <headers.size() ; i++) {
            String header =headers.get(i);
            Map<String,Object> fieldMap =new HashMap<>();
            for (int j = 0; j < list.size(); j++) {
                FieldMsg fieldMsg =list.get(j);
                //没有SET方法跳出当前循环,这个字段不初始化
                if (fieldMsg.getSetMethod()==null){
                    continue;
                }
                Field field =fieldMsg.getField();
                //如果字段名和表头名字一样
                if (field.getName().equals(header)){
                    //如果colIgnore 注解在在字段方法set方法上,或者在字段上忽略此字段
                    if (fieldMsg.getColIgnore()!=null
                            &&("setMethod".equals(fieldMsg.getColIgnoreFrom())
                            || "field".equals(fieldMsg.getColIgnoreFrom()))){
                        break;
                    }
                    fieldMap.put("field",field);
                    fieldMap.put("header",header);
                    list3.add(fieldMap); //其它情况则正面需要给这个字段赋值
                }else {
                    Col col=fieldMsg.getCol();
                    //如果表头和字段名不一样,且没有@col注解，就不知道如何初始化跳过
                    if (col==null){
                        continue;
                    }
                    //如果表头名子注解的名字不一样则这个列不时对应表头跳过
                    if (!col.name().equals(header)){
                        continue;
                    }
                    fieldMap.put("field",field);
                    fieldMap.put("header",header);
                    list3.add(fieldMap);
                    break;
                }
            }
        }
        return list3;
    }
}
