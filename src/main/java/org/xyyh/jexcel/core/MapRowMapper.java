package org.xyyh.jexcel.core;

import org.apache.poi.ss.usermodel.CellType;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.*;

public class MapRowMapper implements RowMapper<LinkedHashMap> {

	private static final Logger logger = LoggerFactory.getLogger(MapRowMapper.class);

	@Override
	public <D> CellValue getData(int rowNum,LinkedHashMap data, int index) {
		// TODO Auto-generated method stub
		CellValue<Object> cellValue = new CellValue<>();
		if(data != null && !data.isEmpty()){
		cellValue.setCellStyle(null);
		cellValue.setCellType(CellType.STRING);
		Set keys= data.keySet();
		if(keys.size()>0){
			List arrayList = new ArrayList<>(keys);
			if(rowNum >0){
				cellValue.setValue(data.get(arrayList.get(index)));
			}else{
				cellValue.setValue(arrayList.get(index));
			}

		}else{
			logger.error("参数map的key为空");
		}
		}else{
			throw new NullPointerException();
		}
		return cellValue;
	}

	@Override
	public int getColumnCount(LinkedHashMap data) {
		// TODO Auto-generated method stub
		int size = data.size();
		return size;
	}

}
