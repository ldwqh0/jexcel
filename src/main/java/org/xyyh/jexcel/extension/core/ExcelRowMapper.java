package org.xyyh.jexcel.extension.core;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

public class ExcelRowMapper{

    private List<String> sheetNames;

    private List<SheetRowMapper> sheetRowMappers;

    private Workbook workbook;

    private int sheets;


    public ExcelRowMapper(String filePath){
        init(filePath);
    }


    public List<String> getSheetNames() {
        return sheetNames;
    }

    public List<SheetRowMapper> getSheetRowMappers() {
        return sheetRowMappers;
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public int getSheets() {
        return sheets;
    }
   public SheetRowMapper  getSheetMapperByNames(String sheetName) {
        for (int i = 0; i <sheets; i++) {
            if (sheetNames.get(i).equals(sheetName)){
                return sheetRowMappers.get(i);
            }
        }
        return null;
    }

   public  void  reLoadByWorkbook(Workbook workbook){
       sheetNames =new ArrayList<>();
       sheetRowMappers =new ArrayList<>();
       int number=workbook.getNumberOfSheets();
       for (int i = 0; i <number ; i++) {
           sheetNames.add(workbook.getSheetAt(i).getSheetName()) ;
           sheetRowMappers.add(new SheetRowMapper(workbook.getSheetAt(i)));
       }
       sheets =sheetNames.size();
   }

    /**
     * @return List<Map<String,Object>>
     *     map.put("header",lisr<string>);
     *     map.put("data",List<LinkedHashMap<String,Object>>);
     *     map.put("name",sheetName);
     *
     * */
    private void init(String filePth){
        toWorkbook(filePth);
        sheetNames =new ArrayList<>();
        sheetRowMappers =new ArrayList<>();
        int number=workbook.getNumberOfSheets();
        for (int i = 0; i <number ; i++) {
            sheetNames.add(workbook.getSheetAt(i).getSheetName()) ;
            sheetRowMappers.add(new SheetRowMapper(workbook.getSheetAt(i)));
        }
        sheets =sheetNames.size();
    }


    /**
     * 根据注解信息排序,顺便吧注解放入map
     * */
    private  void toWorkbook(String filePth){
        if (!StringUtils.isNotBlank(filePth)){
            System.out.println("文件路径错误");
            return ;
        }
        String end =filePth.substring(filePth.lastIndexOf("."));
        if (!end.equalsIgnoreCase(".xls")&&!end.equalsIgnoreCase(".xlsx")){
            System.out.println(".xlsx或者.xls文件");
            return ;
        }
        File file = new File(filePth);
        if (!file.exists()||!file.isFile()){
            System.out.println("文件不存在");
            return ;
        }
        FileInputStream fileInputStream  =null;
        try {

            fileInputStream  = new FileInputStream(file);
            workbook= new XSSFWorkbook(fileInputStream);
        } catch (Exception e) {
            System.out.println("读取文件失败");
            e.printStackTrace();
        }
    }



}
