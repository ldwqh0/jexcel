package org.xyyh.jexcel.utils;

import org.apache.commons.lang3.StringUtils;
import org.apache.log4j.Logger;

import java.math.RoundingMode;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * 日期工具类
 */
public class DateUtis {

	private static final Logger logger = Logger.getLogger(DateUtis.class);
	// 日期类型
	public static int YEAR = 1;
	public static int MONTH = 2;
	public static int WEEK_OF_YEAR = 3;
	public static int WEEK_OF_MONTH = 4;
	public static int DAY_OF_MONTH = 5;
	public static int DAY_OF_YEAR = 6;
	public static int HOUR_OF_DAY = 11;

	/**
	 * 年/月格式化
	 */
	public static final String YEARS = "yyyy";
	public static final String MONTHS = "MM";
	/**
	 * 年-月格式化
	 */
	public static final String YEAR_MONTH = "yyyy-MM";
	/**
	 * 日期 格式化
	 */
	public static final String DATE = "yyyy-MM-dd";
	/**
	 * 日期时间 格式化
	 */
	public static final String TIMESTAMP = "yyyy-MM-dd HH:mm:ss";

	public static final String TIMESTAMPPLUS = "yyyyMMddHHmmss";

	public static final String TIMESTAMPHOUR = "yyyyMMddHH";

	public static final String YYYYMMDD = "yyyyMMdd";
	public static final String MM_DD = "MM-dd";
	public static final String YYYYMM = "yyyyMM";

	/**
	 * 时间 格式化
	 */
	public static final String TIME = "HH:mm:ss";

	/**
	 * 返回指定日期类型日期增量日期
	 *
	 * @param dateType
	 *            日期类型
	 * @param count
	 *            增加数量，可以为负数
	 * @return
	 */
	public static Date getMagicDate(int dateType, int count) {
		Calendar calendar = new GregorianCalendar();
		Date todayTime = new Date();
		calendar.setTime(todayTime);
		calendar.add(dateType, count);
		return calendar.getTime();
	}

	/**
	 * 日期转换为指定格式的字符
	 *
	 * @param date
	 *            日期
	 * @param dateFormat
	 *            日期格式
	 * @return
	 */
	public static String dateToString(Date date, String dateFormat) {
		SimpleDateFormat df = null;
		String returnValue = "";
		if (date != null) {
			df = new SimpleDateFormat(dateFormat);
			returnValue = df.format(date);
		}
		return returnValue;
	}

	/**
	 * 日期字符串转换为日期格式
	 *
	 * @param dateStr
	 *            日期字符串
	 * @param dateFormat
	 *            日期格式
	 * @return
	 * @throws ParseException
	 */
	public static Date stringToDate(String dateStr, String dateFormat) throws ParseException {
		SimpleDateFormat sf = new SimpleDateFormat(dateFormat);
		try {
			return sf.parse(dateStr);
		} catch (ParseException e) {
			logger.error(e);
			throw new ParseException("日期解析错误！", 0);
		}
	}

	// 保留小数点后两位
	public static String formateRate(double rateStr) {
		DecimalFormat f = new DecimalFormat("#.##");
		f.setRoundingMode(RoundingMode.FLOOR);
		return f.format(rateStr);
	}

	/**
	 * 根据毫秒时间获取年
	 *
	 * @param currentTimeMillis
	 * @return
	 */
	public static int getYear(long currentTimeMillis) {
		Calendar calendar = Calendar.getInstance();
		calendar.setTimeInMillis(currentTimeMillis);
		return calendar.get(Calendar.YEAR);
	}

	/**
	 * 根据毫秒时间获取月
	 *
	 * @param currentTimeMillis
	 * @return
	 */
	public static int getMonth(long currentTimeMillis) {
		Calendar calendar = Calendar.getInstance();
		calendar.setTimeInMillis(currentTimeMillis);
		return calendar.get(Calendar.MONTH) + 1;
	}

	public static String getLastMonthStr() {
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.MONTH, -1);
		cal.set(Calendar.DAY_OF_MONTH, 1);
		return new SimpleDateFormat("yyyyMMdd").format(cal.getTime());
	}

	public static String changeFormatToOther(String dateTime, String fromFormat, String toFormat) {
		try {
			return dateToString(stringToDate(dateTime, fromFormat), toFormat);
		} catch (ParseException e) {
			return "";
		}

	}

    /**
     * service返回的数据结构
     * @param success
     * @param msg
     * @param obj
     * @return
     */
	public static Map<String, Object> backResult(String success, String msg, Object obj) {
		final Map<String, Object> result = new HashMap<String, Object>();
		result.put("suc", success);
		result.put("msg", msg);
		result.put("obj", obj);
		return result;
	}

	/**
	 * 获取unix时间戳
	 *
	 * @param t
	 *            时间类型
	 * @param dateFormat
	 *            时间格式，只有在时间格式为字符串时设置
	 * @return long
	 */
	public static <T> long getUnixTime(T t, String dateFormat) {
		long epoch = 0;
		if (null != t) {
			if (t.getClass() == String.class) {
				try {
					Date date = stringToDate((String) t, dateFormat);
					epoch = date.getTime();
				} catch (ParseException e) {
					logger.error(e);
				}
			} else if (t.getClass() == Date.class) {
				epoch = ((Date) t).getTime();
			}
		}
		return epoch;
	}

	/**
	 * 返回指定日期类型日期增量日期
	 *
	 * @param t
	 *            时间类型
	 * @param dateType
	 *            日期类型
	 * @param count
	 *            增加数量，可以为负数
	 * @param dateFormat
	 *            时间格式，只有在时间格式为字符串时设置
	 * @return
	 */
	public static <T> Date getMagicDate(T t, int dateType, int count, String dateFormat) {
		Calendar calendar = new GregorianCalendar();
		Date date = null;
		if (t.getClass() == String.class) {
			// 字符串格式日期
			try {
				date = stringToDate((String) t, dateFormat);
			} catch (ParseException e) {
				logger.error(e);
			}
		} else if (t.getClass() == Date.class) {
			// 日期格式
			date = (Date) t;
		}
		if (null != date) {
			calendar.setTime(date);
			calendar.add(dateType, count);
		}
		return calendar.getTime();
	}

	/**
	 * 2隔字符串时间比较前后
	 *
	 * @param time1
	 *            字符串时间
	 * @param time2
	 *            字符串时间
	 * @param dateFormat
	 *            时间格式
	 * @return int 1 time1大于time2，0相等，-1 time1小于time2
	 */
	public static int compareDate(String time1, String time2, String dateFormat) {
		DateFormat df = new SimpleDateFormat(dateFormat);
		try {
			Date dt1 = df.parse(time1);
			Date dt2 = df.parse(time2);
			if (dt1.getTime() > dt2.getTime()) {
				return 1;
			} else if (dt1.getTime() < dt2.getTime()) {
				return -1;
			} else {
				return 0;
			}
		} catch (Exception e) {
			logger.error(e);
		}
		return 0;
	}

	/**
	 * 获得某个时间区间所有日期
	 *
	 * @param dBeginStr
	 *            开始时间
	 * @param dEndStr
	 *            结束时间
	 * @param pattern
	 *            时间格式
	 * @return
	 * @throws ParseException
	 */
	public static List<String> getBetweenDates(String dBeginStr, String dEndStr, String pattern) throws ParseException {
		pattern = StringUtils.isEmpty(pattern) ? "yyyyMMdd" : pattern;
		SimpleDateFormat sdf = new SimpleDateFormat(pattern);
		List<String> lDate = new ArrayList<String>();
		lDate.add(dBeginStr);
		Date dBegin = sdf.parse(dBeginStr);
		Date dEnd = sdf.parse(dEndStr);
		Calendar calBegin = Calendar.getInstance();
		// 使用给定的 Date 设置此 Calendar 的时间
		calBegin.setTime(dBegin);
		Calendar calEnd = Calendar.getInstance();
		// 使用给定的 Date 设置此 Calendar 的时间
		calEnd.setTime(dEnd);
		// 测试此日期是否在指定日期之后
		while (dEnd.after(calBegin.getTime())) {
			// 根据日历的规则，为给定的日历字段添加或减去指定的时间量
			calBegin.add(Calendar.DAY_OF_MONTH, 1);
			lDate.add(sdf.format(calBegin.getTime()));
		}
		return lDate;
	}

	/**
	 *	获取不同日期相差天数
	 * @param startDay 开始日期，格式"yyyyMMdd"
	 * @param endDay 结束日期，格式"yyyyMMdd"
	 * @return
	 * @throws ParseException
	 */
	public static Long getDaysBetween(String startDay,String endDay) throws ParseException {
		SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd");
		Date d1 = sdf.parse(startDay);
		Date d2 = sdf.parse(endDay);
		long difference=d2.getTime()-d1.getTime();
		long day=difference/(3600*24*1000);
		return day;
	}
	/**
	 * 获取前多天的时间
	 *
	 * @param past
	 * @return
	 */
	public static String getPastDate(int past, String pattern) {
		Calendar calendar = Calendar.getInstance();
		calendar.set(Calendar.DAY_OF_YEAR, calendar.get(Calendar.DAY_OF_YEAR) - past);
		Date today = calendar.getTime();
		SimpleDateFormat format = new SimpleDateFormat(pattern);
		String result = format.format(today);
		return result;
	}

	/**
	 * 获取前几小时的数据
	 *
	 * @param past
	 * @return
	 */
	public static String getPastHour(int past, String pattern) {
		Calendar calendar = Calendar.getInstance();
		calendar.set(Calendar.HOUR_OF_DAY, calendar.get(Calendar.HOUR_OF_DAY) - past);
		Date today = calendar.getTime();
		SimpleDateFormat format = new SimpleDateFormat(pattern);
		String result = format.format(today);
		return result;
	}


	/**
	 *获取当前月第一天和最后一天，并返回
	 * @param pattern
	 * @return
	 */
	public static String getMonthTimeRange(int year,int month,String pattern){
		Calendar cale = Calendar.getInstance();
		SimpleDateFormat format = new SimpleDateFormat(pattern);
		cale.set(Calendar.YEAR,year);
		cale.set(Calendar.MONTH, month);
		cale.set(Calendar.DAY_OF_MONTH, 1);
		cale.add(Calendar.DAY_OF_MONTH, -1);
		String lastday = format.format(cale.getTime());
		cale.set(Calendar.DAY_OF_MONTH, 1);
		String firstday = format.format(cale.getTime());
		String timeRange = firstday + " ~ " + lastday;
		return timeRange;
	}

	public static String getNeedMonth(String currentMonth,String pattern){
		Calendar c = Calendar.getInstance();
		SimpleDateFormat format = new SimpleDateFormat(pattern);
		String month = "";
		try {
			Date date = DateUtis.stringToDate(currentMonth, DateUtis.YYYYMM);
			c.setTime(date);
			c.add(Calendar.MONTH, -1);
			Date m = c.getTime();
			month = format.format(m);
		} catch (ParseException e) {
			e.printStackTrace();
		}
		return month;
	}
}
