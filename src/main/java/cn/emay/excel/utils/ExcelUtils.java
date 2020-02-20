package cn.emay.excel.utils;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Date;

import cn.emay.excel.common.ExcelVersion;

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

	/**
	 * 从路径中解析出版本及输入流
	 * 
	 * @param excelPath
	 *            Exce路径
	 * @return
	 */
	public static ExcelVersion parserPath(String excelPath) {
		ExcelVersion version = null;
		if (excelPath == null) {
			throw new IllegalArgumentException("excelPath is null");
		}
		if (!new File(excelPath).exists()) {
			throw new IllegalArgumentException("excelPath[" + excelPath + "] is not exists");
		}
		if (excelPath.endsWith(ExcelVersion.XLSX.getSuffix())) {
			version = ExcelVersion.XLSX;
		} else if (excelPath.endsWith(ExcelVersion.XLS.getSuffix())) {
			version = ExcelVersion.XLS;
		} else {
			throw new IllegalArgumentException("excelPath[" + excelPath + "] is not excel");
		}
		return version;
	}

}
