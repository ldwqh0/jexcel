package org.xyyh.jexcel.core;

import java.io.*;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.*;

import com.sun.deploy.net.HttpResponse;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import javax.servlet.http.HttpServletResponse;

public class ExcelMapper {

	public <T> Object toExcel(List<T> t) {

		HSSFWorkbook workbook = new HSSFWorkbook();
		return t;
	}

	/**
	 * 由Map输出Excel(制定文件路径)
	 * @param t
	 * @param <T>
	 * @return
	 */
	public <T> Object toExcel(String path, List<LinkedHashMap> t) {
		FileOutputStream fileOut = null;
		try {
			fileOut = new FileOutputStream(path);
			WriteToExcel(t,fileOut);
		} catch (IOException e) {
			e.printStackTrace();
		}finally{
			try {
				fileOut.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return t;
	}

	/**
	 * 实体对象输出excel到制定路径
	 * @param path
	 * @param t
	 * @param <T>
	 * @return
	 */
	public <T> Object toExcelByObject(String path, List<Object> t) {
		FileOutputStream fileOut = null;
		try {
			fileOut = new FileOutputStream(path);
			writeToExcel(t,fileOut);
		} catch (IOException e) {
			e.printStackTrace();
		}finally{
			try {
				fileOut.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return t;
	}

	/**
	 * 由Map通过http从网页导出
	 * @param response
	 * @param t
	 * @param <T>
	 * @return
	 */
	public <T> Object toExcel(String fileName,HttpServletResponse response, List<LinkedHashMap> t) {
		try {
			setHeader(fileName, response);
			WriteToExcel(t,response.getOutputStream());
		} catch (IOException e) {
			e.printStackTrace();
		}finally {
			try {
				response.getOutputStream().close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}

		return t;
	}

	public <T> void writeToExcel(List<T> t, RowMapper<T> rowMapper, OutputStream stream) {
		HSSFWorkbook workbook = new HSSFWorkbook();
		Sheet sheet = workbook.createSheet("sheet1");
		int size = t.size();
		//表格行数
		int rowNum = 0;
		Row row = sheet.createRow(rowNum);
		//写入表头
		writeCells(rowNum,row,t.get(0),rowMapper);
		for (int i = 0; i < size; i++) {
			++rowNum;
			T data = t.get(i);
			row = sheet.createRow(rowNum);
			//逐行写入表单
			writeCells(rowNum,row, data, rowMapper);
		}
		try {
			workbook.write(stream);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/**
	 *	map类型行映射
	 * @param datas
	 * @param stream
	 */
	public void WriteToExcel(List<LinkedHashMap> datas, OutputStream stream) {

		writeToExcel(datas, new MapRowMapper(), stream);
	}

	/**
	 * 简单对象类型行映射
	 * @param datas
	 * @param stream
	 */
	public void writeToExcel(List<Object> datas, OutputStream stream) {
		Object o = datas.get(0);
		Class<?> aClass = o.getClass();
		writeToExcel(datas, new BaseSimpleObjectRowMapper(), stream);
	}

	private <T> void writeCells(int rowNum,Row row, T data, RowMapper<T> rowMapper) {
		int columnSize = rowMapper.getColumnCount(data);

		for (int i = 0; i < columnSize; i++) {
			Cell cell = row.createCell(i);
			CellValue obj = rowMapper.getData(rowNum,data, i);
			writeCell(cell, obj);
		}
	}

	private void writeCell(Cell cell, CellValue cellValue) {
		cell.setCellStyle(cellValue.getCellStyle());
		cell.setCellType(cellValue.getCellType());
		if(cellValue.getValue() instanceof String){
			cell.setCellValue((String) cellValue.getValue());
		}else if(cellValue.getValue() instanceof Number){
			cell.setCellValue( ((Number) cellValue.getValue()).doubleValue());
		}else if(cellValue.getValue() instanceof Date){
			SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
			Date date = (Date) cellValue.getValue();
			cell.setCellValue(sdf.format(date));
		}else if(cellValue.getValue() instanceof Calendar){
			//TODO
			cell.setCellValue((Calendar) cellValue.getValue());
		}else if(cellValue.getValue() instanceof RichTextString){
			//TODO
			cell.setCellValue((RichTextString) cellValue.getValue());
		}else if(cellValue.getValue() instanceof Boolean){
			cell.setCellValue((Boolean) cellValue.getValue());
		}
		System.out.println("每个单元格内容："+cell.toString());
	}

	public <T> List<T> parse(InputStream stream, List<T> datas, Class<T> cls) {
		return datas;
	}

	/**
	 * 设置输出header格式
	 * @param filename
	 * @param response
	 * @throws UnsupportedEncodingException
	 */
	private void setHeader(String filename, HttpServletResponse response) throws UnsupportedEncodingException {
		response.setContentType("application/vnd.ms-excel;charset=UTF-8");
		filename = filename + ".xls";
		String filenameNew = URLEncoder.encode(filename, "UTF-8");
		response.setHeader("accept", "application/vnd.ms-excel");
		response.setHeader("Content-Disposition", "attachment; filename=" + filenameNew);
		response.setHeader("pragma", "no-cache");
		response.setHeader("cache-control", "no-cache");
		response.setHeader("expires", "0");
		response.setCharacterEncoding("UTF-8");
	}
}
