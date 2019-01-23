package org.xyyh.jexcel.extension.core;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xyyh.jexcel.core.ExcelMapper;
import org.xyyh.jexcel.core.RowMapper;

import java.util.ArrayList;
import java.util.List;

import static org.apache.poi.ss.usermodel.CellType.STRING;

public class BeanToExcel {
    private static final Logger logger = LoggerFactory.getLogger(ExcelMapper.class);


    private  Workbook workbook;

    private List<String> sheetNames;

    private List<Sheet> sheets;

    private int sheetNumber;


    public BeanToExcel(){
        workbook =new XSSFWorkbook();
        sheetNames =new ArrayList<>();
        sheets =new ArrayList<>();

    }




    public Workbook getWorkbook() {
        return workbook;
    }

    public List<String> getSheetNames() {
        return sheetNames;
    }

    public List<Sheet> getSheets() {
        return sheets;
    }

    public int getSheetNumber() {
        return sheetNumber;
    }

    public  Sheet getSheetByName(String name) {
        for (int i = 0; i < sheetNames.size(); i++) {
            if (sheetNames.get(i).equals(name)) {
                return sheets.get(i);
            }
        }
     return  null;
    }

    public  void reLoadByWorkbook(Workbook workbook){
        this.workbook =workbook;
        sheetNames =new ArrayList<>();
        sheets =new ArrayList<>();
        int number=workbook.getNumberOfSheets();
        for (int i = 0; i <number ; i++) {
            sheetNames.add(workbook.getSheetAt(i).getSheetName()) ;
            sheets.add(workbook.getSheetAt(i));
        }
        sheetNumber=sheetNames.size();
    }


    /**
     * 指定一组数据和一个{@link RowMapper}来实现数据转换
     *
     * @param datas
     * @param workbook
     * @return
     */
  public  <T> Sheet addSheet(List<T> datas, Workbook workbook ) {
        ObjectRowMapper objectRowMapper =new ObjectRowMapper(datas.get(0).getClass());
        String sheetName =getSheetName(datas.get(0).getClass());
        Sheet sheet1 = workbook.createSheet(sheetName);
         //初始化表头
        addHeader(sheet1,objectRowMapper);
        for (int rowIndex = 1; rowIndex <= datas.size(); rowIndex++) {
            addRowAndSetValue(sheet1,datas.get(rowIndex-1),objectRowMapper,rowIndex);
        }
        sheets.add(sheet1);
        sheetNames.add(sheetName);
        sheetNumber =sheetNames.size();
        return sheet1;
    }

  private static String   getSheetName(Class tClass){
        org.xyyh.jexcel.annotations.Sheet annotation =
                      (org.xyyh.jexcel.annotations.Sheet) tClass.getAnnotation(org.xyyh.jexcel.annotations.Sheet.class);
              if (annotation==null){
                  return tClass.getSimpleName();
              }else if ("".equals(annotation.name())){
                  return tClass.getSimpleName();
              }else {
                  return annotation.name();
              }
        }
    private static<T> void  addRowAndSetValue(Sheet sheet,T bean,RowMapper<T> rowMapper,int rowIndex){
           Row row = sheet.createRow(rowIndex);
           for (int i = 0; i < rowMapper.getColumnCount(); i++) {
               CellValue  cellValue = rowMapper.getCellValue(i,bean);
               Cell cell = row.createCell(i);
               writeCell(cell,cellValue);
           }
     }

    private static<T> void  addHeader(Sheet sheet,RowMapper<T> rowMapper){
        Row row = sheet.createRow(0);
         List<String> list = rowMapper.getHeaders();
        for (int i = 0; i < list.size(); i++) {
            Cell cell = row.createCell(i,STRING);
            cell.setCellValue(list.get(i));
        }
    }
    /**
     * 写入单元格值
     *
     * @param cell
     * @param cellValue
     */
    private static void writeCell(Cell cell, CellValue cellValue) {
        cell.setCellType(cellValue.getCellType());
        switch (cellValue.getCellType()) {
            case NUMERIC:
                cell.setCellValue(cellValue.getNumberValue());
                break;
            case BOOLEAN:
                cell.setCellValue(cellValue.getBooleanValue());
                break;
            case STRING:
                cell.setCellValue(cellValue.getStringValue());
                break;
            case FORMULA: // 暂时不处理公式
                break;
            case _NONE: // 对于空值不处理
                break;
            case BLANK:
                break;
            case ERROR:
                break;
            default:
        }
    }
}
