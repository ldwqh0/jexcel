package org.xyyh.jexcel.core;

import java.io.*;
import java.lang.reflect.Field;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xyyh.jexcel.utils.DateUtis;
import org.xyyh.jexcel.vo.FieldForSortting;

import static org.apache.poi.ss.usermodel.CellType.*;

public class ExcelMapper {

	private static final Logger logger = LoggerFactory.getLogger(ExcelMapper.class);


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
	}

	/**
	 * 解析excel,并做导入
	 * @param stream
	 * @param datas
	 * @param cls
	 * @param <T>
	 * @return
	 */
	public <T> List<T> parse(InputStream stream, List<T> datas, Class<T> cls) {
		Workbook workbook = null;
		try {
			workbook = WorkbookFactory.create(stream);
		} catch (IOException e) {
			//获取工作簿失败
			logger.error("load excel file error: ",e);
		}
		Sheet sheet;
		Iterator<Row> rowIterator = null;
		if (workbook != null) {
			sheet = workbook.getSheetAt(0);
			rowIterator = sheet.rowIterator();
		}

		try{
			//表头与位置对照
			Map<String,Integer> titleMap = new HashMap<>();
			while(true){
				Row row;
				do {
					while(rowIterator.hasNext()){
						row = rowIterator.next();
						//略过表头
						if(row.getRowNum() == 0){
							Integer sortNum = 0;
							//titlemap填充
							Iterator<Cell> cellIterator1 = row.cellIterator();
							while(cellIterator1.hasNext()){
								String key = cellIterator1.next().toString();
								titleMap.put(key,sortNum);
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
							//输出map
							if(cls == Map.class){
								Map<String,Object> map = new HashMap<>();
								//获取属性和位置对照，并写入map
								for (Object o : titleMap.keySet()) {
									String key = (String) o;
									Integer index = (Integer) titleMap.get(key);
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
							}else{
								T t = cls.newInstance();
								//实体类属性字段编号
								List<FieldForSortting> fileds = sortFieldByAnno(cls);
								Iterator<FieldForSortting> filedIterator = fileds.iterator();
								while(true){
									while(filedIterator.hasNext()){
										FieldForSortting ffs =  filedIterator.next();
										Field field = ffs.getField();
										String fieldName = ffs.getFieldName();
										field.setAccessible(true);
										//对应位置单元格内容
										Integer sortNum = titleMap.get(fieldName);
										if(sortNum == null){
											break;
										}
										Cell cell = row.getCell(sortNum);
										Object cellValue = getCellValue(cell);
										//属性类型转换
										if(field.getType().equals(Date.class)){
											Date value = DateUtis.stringToDate((String) cellValue, DateUtis.TIMESTAMP);
											field.set(t, value);
										}else if(field.getType().equals(Integer.class)){
											Integer value = ((Double) cellValue).intValue();
											field.set(t, value);
										}else if (field.getType().equals(Long.class)){
											Long value = ((Double) cellValue).longValue();
											field.set(t, value);
										}else if(field.getType().equals(Byte.class)){
											byte value = ((Double) cellValue).byteValue();
											field.set(t, value);
										}else if(field.getType().equals(Short.class)){
											short value = ((Double) cellValue).shortValue();
											field.set(t,value);
										}else if(field.getType().equals(Float.class)){
											float value = ((Double) cellValue).floatValue();
											field.set(t,value);
										}else{
											field.set(t, cellValue);
										}

									}
									datas.add(t);
									break;
								}
							}
						}

					}
					return datas;
				}while(cls != Map.class);

			}
		} catch (IllegalAccessException
				| InstantiationException
				| ParseException e) {
			e.printStackTrace();
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

	/**
	 * 获取单元格的值
	 * @param cell
	 * @return
	 */
	private Object getCellValue(Cell cell) {
		if (cell != null && (cell.getCellType() != CellType.STRING || !StringUtils.isBlank(cell.getStringCellValue()))) {
			CellType cellType = cell.getCellType();
			if (cellType == CellType.BLANK) {
				return null;
			} else if (cellType == CellType.BOOLEAN) {
				return cell.getBooleanCellValue();
			} else if (cellType == CellType.ERROR) {
				return cell.getErrorCellValue();
			} else if (cellType == CellType.FORMULA) {
				/*try {
					return HSSFDateUtil.isCellDateFormatted(cell) ? cell.getDateCellValue() : cell.getNumericCellValue();
				} catch (IllegalStateException var3) {
					return cell.getRichStringCellValue();
				}*/
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
	 * 给实体类属性值制定顺序
	 * @param clazz
	 * @return
	 */
	private  List<FieldForSortting> sortFieldByAnno(Class<?> clazz)  {
		List<FieldForSortting> list = new ArrayList<>();
		getAllFields(list,clazz);
		return list;
	}

	/**
	 * 递归调用，获取本类所有属性
	 * @param list
	 * @param clazz
	 */
	private void getAllFields(List<FieldForSortting> list, Class<?> clazz){
		if(clazz != Object.class){
			returnClassField(list, clazz);
			Class supperClazz = clazz.getSuperclass();
			getAllFields(list,supperClazz);
		}
	}

	/**
	 * 获取当前类所有属性
	 * @param list
	 * @param clazz
	 */
	private void returnClassField(List<FieldForSortting> list, Class<?> clazz) {
		Field[] declaredFields = clazz.getDeclaredFields();
		for (Field declaredField : declaredFields) {
			FieldForSortting ffs = new FieldForSortting();
			declaredField.setAccessible(true);
			ffs.setField(declaredField);
			ffs.setFieldName(declaredField.getName());
			ffs.setIndex(0);
			list.add(ffs);
		}
	}
}
