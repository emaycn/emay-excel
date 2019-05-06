package cn.emay.excel.utils;

import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 一些工具
 * 
 * @author Frank
 *
 */
public class ExcelUtils {

	/**
	 * 获取int类型的表达式
	 * 
	 * @return
	 */
	public static int parserExpressToInt(String express) {
		int num = -1;
		if (express != null && !"".equalsIgnoreCase(express.trim())) {
			try {
				return Integer.parseInt(express);
			} catch (Exception e) {
			}
		}
		return num;
	}

	/**
	 * 把字符串转成日期
	 * 
	 * @param dateStr
	 *            日期字符串
	 * @param format
	 *            日期格式
	 * @return
	 */
	public static Date parseDate(String dateStr, String format) {
		Date date = null;
		try {
			SimpleDateFormat sdf = new SimpleDateFormat(format);
			date = sdf.parse(dateStr);
		} catch (Exception e) {
		}
		return date;
	}

	/**
	 * 新建一个数据实例
	 * 
	 * @return
	 */
	public static <D> D newData(Class<D> dataClass) {
		try {
			return dataClass.newInstance();
		} catch (InstantiationException e) {
			throw new IllegalArgumentException(dataClass.getName() + " can't be new Instance", e);
		} catch (IllegalAccessException e) {
			throw new IllegalArgumentException(dataClass.getName() + " can't be new Instance", e);
		}
	}

}
