package org.xyyh.jexcel.core;

import java.io.*;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.*;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.collections4.IterableUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import static org.apache.poi.ss.usermodel.CellType.*;

public class ExcelMapper {

	public <T> Object toExcel(List<T> t) {
		HSSFWorkbook workbook = new HSSFWorkbook();
		return t;
	}

	/**
	 * 由Map输出Excel(制定文件路径)
	 * 
	 * @param t
	 * @param   <T>
	 * @return
	 */
	public <T> Object toExcel(String path, List<LinkedHashMap> t) {
		FileOutputStream fileOut = null;
		try {
			fileOut = new FileOutputStream(path);
			WriteToExcel(t, fileOut);
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
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
	 * 
	 * @param path
	 * @param t
	 * @param      <T>
	 * @return
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	public void toExcelByObject(String path, List<Object> t) throws FileNotFoundException, IOException {
		try (FileOutputStream fileOut = new FileOutputStream(path)) {
//			fileOut
//			writeToExcel(t, fileOut);
		}
//		return t;
	}

	/**
	 * 由Map通过http从网页导出
	 * 
	 * @param response
	 * @param t
	 * @param          <T>
	 * @return
	 */
	public <T> Object toExcel(String fileName, HttpServletResponse response, List<LinkedHashMap> t) {
		try {
			setHeader(fileName, response);
			WriteToExcel(t, response.getOutputStream());
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
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
		// 表格行数
		int rowNum = 0;
		Row row = sheet.createRow(rowNum);
		// 写入表头
		writeCells(rowNum, row, t.get(0), rowMapper);
		for (int i = 0; i < size; i++) {
			++rowNum;
			T data = t.get(i);
			row = sheet.createRow(rowNum);
			// 逐行写入表单
			writeCells(rowNum, row, data, rowMapper);
		}
		try {
			workbook.write(stream);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/**
	 * map类型行映射
	 * 
	 * @param datas
	 * @param stream
	 */
	public void WriteToExcel(List<LinkedHashMap> datas, OutputStream stream) {
//		writeToExcel(datas, new MapRowMapper(), stream);
	}

	/**
	 * 简单对象类型行映射
	 * 
	 * @param datas
	 * @param stream
	 */
	public <T> void writeToExcel(OutputStream stream, List<T> datas, Class<T> tClass) {
		Object o = datas.get(0);
//		tClass.get
//		writeToExcel(datas, new BaseSimpleObjectRowMapper(), stream);
	}

	private <T> void writeCells(int rowNum, Row row, T data, RowMapper<T> rowMapper) {
		int columnSize = rowMapper.getColumnCount(data);

		for (int i = 0; i < columnSize; i++) {
			Cell cell = row.createCell(i);
			SimpleCellValue obj = rowMapper.getData(rowNum, data, i);
			writeCell(cell, obj);
		}
	}

	private void writeCell(Cell cell, SimpleCellValue cellValue) {
		cell.setCellStyle(cellValue.getCellStyle());
		cell.setCellType(cellValue.getCellType());
		if (cellValue.getValue() instanceof String) {
			cell.setCellValue((String) cellValue.getValue());
		} else if (cellValue.getValue() instanceof Number) {
			cell.setCellValue(((Number) cellValue.getValue()).doubleValue());
		} else if (cellValue.getValue() instanceof Date) {
			SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
			Date date = (Date) cellValue.getValue();
			cell.setCellValue(sdf.format(date));
		} else if (cellValue.getValue() instanceof Calendar) {
			// TODO
			cell.setCellValue((Calendar) cellValue.getValue());
		} else if (cellValue.getValue() instanceof RichTextString) {
			// TODO
			cell.setCellValue((RichTextString) cellValue.getValue());
		} else if (cellValue.getValue() instanceof Boolean) {
			cell.setCellValue((Boolean) cellValue.getValue());
		}
		System.out.println("每个单元格内容：" + cell.toString());
	}

	public <T> List<T> parse(InputStream stream, List<T> datas, Class<T> cls) {
		return datas;
	}

	/**
	 * 设置输出header格式 暂时不要关心下载的问题
	 * 
	 * @param filename
	 * @param response
	 * @throws UnsupportedEncodingException
	 */
	@Deprecated
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

	/**
	 * 将excel文件保存到指定的输出流
	 * 
	 * @param datas
	 * @param out
	 * @throws IOException
	 */
	public <T> void toExcel(List<T> datas, OutputStream out) throws IOException {
		try (Workbook workBook = toWorkBook(datas)) {
			workBook.write(out);
		}
	}

	/**
	 * 将excel文件保存到指定的文件
	 * 
	 * @param datas
	 * @param out
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public <T> void toExcel(List<T> datas, File out) throws FileNotFoundException, IOException {
		try (FileOutputStream outStream = new FileOutputStream(out)) {
			toExcel(datas, outStream);
		}
	}

	/**
	 * 将excel保存到指定
	 * 
	 * @param datas
	 * @param path
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public <T> void toExcel(List<T> datas, String path) throws FileNotFoundException, IOException {
		try (FileOutputStream outStream = new FileOutputStream(path)) {
			toExcel(datas, outStream);
		}
	}

	/**
	 * 将一组{@link Map}转换为一个工作簿
	 * 
	 * @param datas
	 * @return
	 */
	private Workbook mapToWorkBook(List<Map<String, Object>> datas) {
		if (CollectionUtils.isNotEmpty(datas)) {
			Map<String, Object> sample = datas.get(0);
			return this.toWorkBook(datas, new AutoColumnRowMapper(sample));
		} else {
			return this.toWorkBook(Collections.<Map<String, Object>>emptyList(),
					Collections.<String>emptyList(),
					Collections.<String>emptyList());
		}
	}

	/**
	 * 将一组对象转换为Excel表格
	 * @param datas
	 * @return
	 */
	private <T> Workbook objectToWorkBook(List<T> datas) {
		// TODO 将一组对象转换为workbook
		return null;
	}

	/**
	 * 将一组map对象转换为excel对象 <br />
	 * 请给定一组键的序列，以便确定导出数据的序列
	 * 
	 * @author LiDong
	 * @param datas   要转换的对象
	 * @param headers 转换的对象的表头值
	 * @param keys    转换的对象的键序列
	 * @return
	 */
	public Workbook toWorkBook(List<Map<String, Object>> datas, List<String> headers, List<String> keys) {
		RowMapper<Map<String, Object>> rowMapper = new FixedColumnMapRowMapper(headers, keys);
		return toWorkBook(datas, rowMapper);
	}

	/**
	 * 指定一组数据和一个{@link RowMapper}来实现数据转换
	 * 
	 * @param datas
	 * @param rowMapper
	 * @return
	 */
	public <T> Workbook toWorkBook(List<T> datas, RowMapper<T> rowMapper) {
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet1 = workbook.createSheet();
		List<String> header = rowMapper.getHeaders();
		Row headerRow = sheet1.createRow(0);
		int columnCount = rowMapper.getColumnCount();
		for (int i = 0; i < columnCount; i++) {
			Cell cell = headerRow.createCell(i, STRING);
			cell.setCellValue(header.get(i));
		}
		int size = datas.size();
		for (int rowIndex = 1; rowIndex <= size; rowIndex++) {
			int dataIndex = rowIndex - 1;
			T data = datas.get(dataIndex);
			Row dataRow = sheet1.createRow(rowIndex);
			for (int colIndex = 0; colIndex < columnCount; colIndex++) {
				CellValue cellValue = rowMapper.getCellValue(colIndex, data);
				Cell cell = dataRow.createCell(colIndex);
				writeCell(cell, cellValue);
			}
		}
		return workbook;
	}

	/**
	 * 将一组数据转换为excel表格
	 * 
	 * @param datas
	 * @return
	 */
	@SuppressWarnings("unchecked")
	public <T> Workbook toWorkBook(List<T> datas) {
		if (CollectionUtils.isNotEmpty(datas)) {
			T data = datas.get(0);
			if (data instanceof Map) {
				return mapToWorkBook((List<Map<String, Object>>) datas);
			} else {
				return objectToWorkBook(datas);
			}
		} else {
			return null;
		}
	}

	/**
	 * 写入单元格值
	 * 
	 * @param cell
	 * @param cellValue
	 */
	private void writeCell(Cell cell, CellValue cellValue) {
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
		}
	}
}
