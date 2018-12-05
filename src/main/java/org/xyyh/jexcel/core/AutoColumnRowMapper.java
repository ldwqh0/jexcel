package org.xyyh.jexcel.core;

import java.util.Collections;
import java.util.List;
import java.util.Map;

import org.apache.commons.collections4.IterableUtils;
import org.apache.commons.collections4.map.LinkedMap;

public class AutoColumnRowMapper extends FixedColumnMapRowMapper {
	public AutoColumnRowMapper(Map<String, Object> sample) {
		super();
		List<String> keys = IterableUtils.toList(sample.keySet());
		if (sample instanceof LinkedMap || sample instanceof LinkedMap) {
			super.setHeaders(keys);
			super.setHeaders(keys);
		} else {
			Collections.sort(keys);
		}
		setHeaders(keys);
		setKeys(keys);
	}
}
