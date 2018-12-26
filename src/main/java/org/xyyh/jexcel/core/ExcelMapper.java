package org.xyyh.jexcel.core;

import java.io.*;
import java.lang.reflect.Field;
import java.text.ParseException;
import java.util.*;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xyyh.jexcel.utils.DateUtis;
import org.xyyh.jexcel.utils.FieldUtils;
import org.xyyh.jexcel.vo.FieldForSortting;

import static org.apache.poi.ss.usermodel.CellType.*;

public class ExcelMapper {

	private static final Logger logger = LoggerFactory.getLogger(ExcelMapper.class);

	/**
	 * 解析excel,并做导入
	 * 
	 * @param stream
	 * @param datas
	 * @param cls
	 * @param        <T>
	 * @return
	 */
	public <T> List<T> parse(InputStream stream, List<T> datas, Class<T> cls) throws IllegalAccessException,
			ParseException, InstantiationException {
		Workbook workBook = getWorkBook(stream);
		if (workBook != null) {
			Sheet sheet = workBook.getSheetAt(0);
			Iterator<Row> rowIterator = sheet.rowIterator();
			//返回集合信息
			datas = workBookToList(rowIterator, datas, cls);
		} else {
			logger.error("Failed to get the file, please check if the file format is correct");
		}
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
		Sheet sheet1 = null;
		Class<?> aClass = datas.get(0).getClass();
		boolean isPresent = aClass.isAnnotationPresent(org.xyyh.jexcel.annotations.Sheet.class);
		if(isPresent){
			org.xyyh.jexcel.annotations.Sheet annotation =
					aClass.getAnnotation(org.xyyh.jexcel.annotations.Sheet.class);
			String name = annotation.name();
			if (!"".equals(name)){
				sheet1 = workbook.createSheet(name);
			}
		}else{
			sheet1 = workbook.createSheet();
		}
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
		default:

		}
	}

	/**
	 * 获取单元格的值
	 * 
	 * @param cell
	 * @return
	 */
	private Object getCellValue(Cell cell) {
		if (cell != null
				&& (cell.getCellType() != CellType.STRING || !StringUtils.isBlank(cell.getStringCellValue()))) {
			CellType cellType = cell.getCellType();
			if (cellType == CellType.BLANK) {
				return null;
			} else if (cellType == CellType.BOOLEAN) {
				return cell.getBooleanCellValue();
			} else if (cellType == CellType.ERROR) {
				return cell.getErrorCellValue();
			} else if (cellType == CellType.FORMULA) {
				return null;
			} else if (cellType == CellType.NUMERIC) {
				return DateUtil.isCellDateFormatted(cell) ? cell.getDateCellValue() : cell.getNumericCellValue();
			} else {
				return cellType == CellType.STRING ? cell.getStringCellValue() : null;
			}
		} else {
			return null;
		}
	}

	/**
	 * 获取工作簿对象
	 * @param stream
	 * @return
	 */
	private Workbook getWorkBook(InputStream stream) {
		//创建Workbook工作薄对象，表示整个excel
		Workbook workbook = null;
		try {
			//2007+
			workbook = new XSSFWorkbook(stream);
		} catch (IOException e) {
			logger.error("load excel file error: ",e);
		}
		return workbook;
	}

	/**
	 * 获取表头信息，按转换类型不同分别转换
	 * @param rowIterator
	 * @param datas
	 * @param cls
	 * @param <T>
	 * @return
	 * @throws IllegalAccessException
	 * @throws ParseException
	 * @throws InstantiationException
	 */
	private <T> List<T> workBookToList(Iterator<Row> rowIterator,List<T> datas, Class<T> cls) throws IllegalAccessException, ParseException, InstantiationException {
		Map<String, Integer> titleMap = new HashMap<>();
		Row row;
		while (rowIterator.hasNext()) {
			row = rowIterator.next();
			// 获取表头，并写入map
			if (row.getRowNum() == 0) {
				Integer sortNum = 0;
				// titleMap填充
				Iterator<Cell> cellIterator = row.cellIterator();
				while (cellIterator.hasNext()) {
					String key = cellIterator.next().toString();
					titleMap.put(key, sortNum);
					sortNum++;
				}
				continue;
			}
			//判断一行是否为空
			boolean allRowIsNull = true;
			Iterator<Cell> cellIterator = row.cellIterator();
			while(cellIterator.hasNext()){
				Object cellValue = getCellValue( cellIterator.next());
				if(cellValue != null){
					allRowIsNull = false;
					break;
				}
			}
			if (allRowIsNull){
				logger.error("Excel row " + row.getRowNum() + " all row value is null!");
			}else{
				if (cls == Map.class){
					workBookToMapList(datas,row,titleMap);
				}else{
					workBookToObjectList(datas,row,titleMap,cls);
				}
			}
		}
		return datas;
	}

	/**
	 *	解析excel生成map集合
	 * @param datas
	 * @param row
	 * @param titleMap
	 * @param <T>
	 */
	@SuppressWarnings("unchecked")
	private <T> void workBookToMapList(List<T> datas, Row row, Map<String, Integer> titleMap) {
		Map<String,Object> map = new HashMap<>();
		//获取属性和位置对照，并写入map
		for (Object o : titleMap.keySet()) {
			String key = (String) o;
			Integer index = titleMap.get(key);
			Cell cell = row.getCell(index);
			if (cell == null) {
				map.put(key, null);
			} else {
				cell.setCellType(CellType.STRING);
				String value = cell.getStringCellValue();
				map.put(key, value);
			}
		}
		datas.add((T) map);
	}

	/**
	 * 解析excel生成实体对象集合
	 * @param datas 实体对象集合
	 * @param row excel的行
	 * @param titleMap 表头名称与位置
	 * @param cls 类属性
	 * @param <T>  泛型标志
	 * @throws IllegalAccessException
	 * @throws InstantiationException
	 * @throws ParseException
	 */
	private <T> void workBookToObjectList(List<T> datas, Row row, Map<String, Integer> titleMap, Class<T> cls) throws IllegalAccessException, InstantiationException, ParseException {
		T t = cls.newInstance();
		// 实体类属性字段编号
		List<FieldForSortting> fields = FieldUtils.sortFieldByAnno(cls);
		for (FieldForSortting ffs : fields) {
			Field field = ffs.getField();
			String fieldName = ffs.getFieldName();
			field.setAccessible(true);
			// 对应位置单元格内容
			Integer sortNum = titleMap.get(fieldName);
			if (sortNum == null) {
				break;
			}
			Cell cell = row.getCell(sortNum);
			Object cellValue = getCellValue(cell);
			// 属性类型转换
			if (field.getType().equals(Date.class)) {
				Date value = DateUtis.stringToDate((String) cellValue, DateUtis.TIMESTAMP);
				field.set(t, value);
			} else if (field.getType().equals(Integer.class)) {
				Integer value = ((Double) cellValue).intValue();
				field.set(t, value);
			} else if (field.getType().equals(Long.class)) {
				Long value = ((Double) cellValue).longValue();
				field.set(t, value);
			} else if (field.getType().equals(Byte.class)) {
				byte value = ((Double) cellValue).byteValue();
				field.set(t, value);
			} else if (field.getType().equals(Short.class)) {
				short value = ((Double) cellValue).shortValue();
				field.set(t, value);
			} else if (field.getType().equals(Float.class)) {
				float value = ((Double) cellValue).floatValue();
				field.set(t, value);
			} else {
				field.set(t, cellValue);
			}
		}
		datas.add(t);
	}

}
