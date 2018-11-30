package org.xyyh.jexcel.core;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelMapper {

	public <T> Object toExcel(List<T> t) {

		HSSFWorkbook workbook = new HSSFWorkbook();

		return t;
	}

	public <T> void writeToExcel(List<T> t, RowMapper<T> rowMapper, OutputStream stream) {
		Workbook workbook = new HSSFWorkbook();
		Sheet sheet = workbook.createSheet("sheet1");
		int size = t.size();
		for (int i = 0; i < size; i++) {
			T data = t.get(i);
			Row row = sheet.createRow(i);
			writeCells(row, data, rowMapper);
		}
	}

	private <T> void writeCells(Row row, T data, RowMapper<T> rowMapper) {
		int columnSize = rowMapper.getColumnCount(data);
		for (int i = 0; i < columnSize; i++) {
			Cell cell = row.createCell(i);
			CellValue obj = rowMapper.getData(data, i);
			writeCell(cell, obj);
		}
	}

	private void writeCell(Cell cell, CellValue cellValue) {
		cell.setCellStyle(cellValue.getCellStyle());
		cell.setCellType(cellValue.getCellType());
		cell.setCellValue((double) cellValue.getValue());
	}

	public <T> List<T> parse(InputStream stream, List<T> datas, Class<T> cls) {
		return datas;
	}

}
