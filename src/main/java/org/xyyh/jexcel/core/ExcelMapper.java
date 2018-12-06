package org.xyyh.jexcel.core;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static org.apache.poi.ss.usermodel.CellType.*;

public class ExcelMapper {

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
	 * 
	 * @param datas
	 * @return
	 */
	@SuppressWarnings("unchecked")
	private <T> Workbook objectToWorkBook(List<T> datas) {
		if (CollectionUtils.isNotEmpty(datas)) {
			T data = datas.get(0);
			return toWorkBook(datas, new ObjectRowMapper<T>((Class<T>) data.getClass()));
		} else {
			return null;
		}
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
