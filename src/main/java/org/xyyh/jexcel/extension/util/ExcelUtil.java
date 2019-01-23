package org.xyyh.jexcel.extension.util;

import org.apache.poi.ss.usermodel.Workbook;
import org.xyyh.jexcel.extension.core.BeanToExcel;
import org.xyyh.jexcel.extension.core.ExcelRowMapper;
import org.xyyh.jexcel.extension.core.MyUserDetail;
import org.xyyh.jexcel.extension.core.SheetRowMapper;

import java.io.File;
import java.io.FileOutputStream;
import java.util.*;

public class ExcelUtil {

    public static void main(String[] args) {
        List<MyUserDetail> myUserDetails =new ArrayList<>();
        for (int i = 0; i < 100000; i++) {
            MyUserDetail myUserDetail =new MyUserDetail();
            myUserDetail.setDate(new Date());
            myUserDetail.setClientSecret("sha");
            myUserDetail.setName("shei");
            myUserDetail.setClientId("zhasdiasndyhqw");
            myUserDetails.add(myUserDetail);
        }
        String url ="C:\\Users\\Administrator\\Desktop\\test.xls";
        System.out.println("开始写如100000条数据");
        long b =System.currentTimeMillis();
        BeanToExcel beanToExcel = getEmptyBeanToExcel();
        beanToExcel.addSheet(myUserDetails,beanToExcel.getWorkbook());
        writeExel(beanToExcel.getWorkbook(),url);
        System.out.println("写如数据完成，耗时："+((System.currentTimeMillis()-b)/1000)+"秒");

        b =System.currentTimeMillis();
        System.out.println("开始写如100000条数据");
        ExcelRowMapper excelRowMapper = readExcel(url);
        Map<String,List> map =new HashMap<>();
        SheetRowMapper sheetRowMapper = excelRowMapper.getSheetMapperByNames("MyUserDetail");
        List list =sheetRowMapper.getBeanist(Map.class);
        map.put(sheetRowMapper.getSheetName(),list);
        System.out.println("读取数据，耗时："+((System.currentTimeMillis()-b)/1000)+"秒");
        System.out.println(map.toString());


        //映射间的相互转换




    }
    /**
     * 获取excel的映射实体
     * @param  filePath 文件路径
     * */
    public static ExcelRowMapper readExcel(String filePath){
        return  new ExcelRowMapper(filePath);
    }

    /**
     * 获取一个空的BeanToExcel
     * */
    public  static void writeExel(Workbook workbook,String filePath){
        File file =new File(filePath);
        /*if (!file.exists()||!file.isFile()){
            System.out.println("文件路径不存在");
            return;
        }*/
        try {
            FileOutputStream fileOutputStream =new FileOutputStream(file);
            workbook.write(fileOutputStream);
            workbook.close();
        } catch (Exception e) {
            System.out.println("写入文件出错");
            e.printStackTrace();

        }
    }
    public static BeanToExcel getEmptyBeanToExcel( ){
        return  new BeanToExcel();
    }


}
