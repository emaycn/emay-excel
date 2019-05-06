package cn.emay.excel.utils;

import java.math.BigDecimal;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

/**
 * 读工具类
 * 
 * @author Frank
 *
 */
public class ExcelReadUtils {

	/**
	 * 读取日期类型数据
	 * 
	 * @param cell
	 *            单元格
	 * @param format
	 *            日期格式
	 * @return [可能为空]
	 */
	public static Date readDate(Cell cell, String express) {
		if (cell == null) {
			return null;
		}
		CellType ctype = cell.getCellTypeEnum();
		Date date = null;
		switch (ctype) {
		case NUMERIC:
			date = cell.getDateCellValue();
			break;
		case STRING:
			date = ExcelUtils.parseDate(cell.getStringCellValue(), express);
			break;
		default:
			break;
		}
		return date;
	}

	/**
	 * 读取日期类型数据
	 * 
	 * @param formatIndex
	 *            单元格类型
	 * @param value
	 *            数据
	 * @param format
	 *            日期格式
	 * @return [可能为空]
	 */
	public static Date readDate(int formatIndex, String value, String express) {
		if (value == null) {
			return null;
		}
		Date date = null;
		try {
			if (formatIndex == 49 || formatIndex == -1) {
				date = ExcelUtils.parseDate(value, express);
			} else {
				Double d1 = Double.valueOf(value);
				date = HSSFDateUtil.getJavaDate(d1);
			}
		} catch (Exception e) {
		}
		return date;
	}

	/**
	 * 读取Integer类型数据
	 * 
	 * @param cell
	 *            单元格
	 * @return [可能为空]
	 */
	public static Integer readInteger(Cell cell) {
		BigDecimal lo = readBigDecimal(cell, -1);
		if (lo != null) {
			return lo.intValue();
		} else {
			return null;
		}
	}

	/**
	 * 读取Integer类型数据
	 * 
	 * @param value
	 *            数据
	 * @return [可能为空]
	 */
	public static Integer readInteger(String value) {
		BigDecimal lo = readBigDecimal(value, -1);
		if (lo != null) {
			return lo.intValue();
		} else {
			return null;
		}
	}

	/**
	 * 读取Long类型数据
	 * 
	 * @param cell
	 *            单元格
	 * @return [可能为空]
	 */
	public static Long readLong(Cell cell) {
		BigDecimal lo = readBigDecimal(cell, -1);
		if (lo != null) {
			return lo.longValue();
		} else {
			return null;
		}
	}

	/**
	 * 读取Long类型数据
	 * 
	 * @param value
	 *            数据
	 * @return [可能为空]
	 */
	public static Long readLong(String value) {
		BigDecimal lo = readBigDecimal(value, -1);
		if (lo != null) {
			return lo.longValue();
		} else {
			return null;
		}
	}

	/**
	 * 读取Double类型数据
	 * 
	 * @param cell
	 *            单元格
	 * @param scale
	 *            保留小数点后位数。(<0则不改变原有值)
	 * @return [可能为空]
	 */
	public static Double readDouble(Cell cell, int scale) {
		BigDecimal lo = readBigDecimal(cell, scale);
		if (lo != null) {
			return lo.doubleValue();
		} else {
			return null;
		}
	}

	/**
	 * 读取Double类型数据
	 * 
	 * @param value
	 *            数据
	 * @param scale
	 *            保留小数点后位数。(<0则不改变原有值)
	 * @return [可能为空]
	 */
	public static Double readDouble(String value, int scale) {
		BigDecimal lo = readBigDecimal(value, scale);
		if (lo != null) {
			return lo.doubleValue();
		} else {
			return null;
		}
	}

	/**
	 * 读取BigDecimal类型数据
	 * 
	 * @param cell
	 *            单元格
	 * @param scale
	 *            保留小数点后位数。(<0则不改变原有值)
	 * @return [可能为空]
	 */
	public static BigDecimal readBigDecimal(Cell cell, int scale) {
		if (cell == null) {
			return null;
		}
		CellType ctype = cell.getCellTypeEnum();
		BigDecimal d2 = null;
		switch (ctype) {
		case NUMERIC:
			d2 = new BigDecimal(cell.getNumericCellValue());
			break;
		case STRING:
			try {
				d2 = new BigDecimal(cell.getStringCellValue());
			} catch (Exception e) {
			}
			break;
		case BOOLEAN:
			d2 = new BigDecimal(cell.getBooleanCellValue() ? 1d : 0d);
			break;
		default:
			break;
		}
		if (d2 != null && scale >= 0) {
			d2 = d2.setScale(scale, BigDecimal.ROUND_HALF_UP);
		}
		return d2;
	}

	/**
	 * 读取BigDecimal类型数据
	 * 
	 * @param value
	 *            数据
	 * @param scale
	 *            保留小数点后位数。(<0则不改变原有值)
	 * @return [可能为空]
	 */
	public static BigDecimal readBigDecimal(String value, int scale) {
		if (value == null) {
			return null;
		}
		BigDecimal d2 = null;
		try {
			if (scale >= 0) {
				d2 = new BigDecimal(value).setScale(scale, BigDecimal.ROUND_HALF_UP);
			} else {
				d2 = new BigDecimal(value);
			}
		} catch (Exception e) {
		}
		return d2;
	}

	/**
	 * 读取String类型数据
	 * 
	 * @param cell
	 *            单元格
	 * @return [可能为空]
	 */
	public static String readString(Cell cell) {
		if (cell == null) {
			return null;
		}
		CellType ctype = cell.getCellTypeEnum();
		String str = null;
		switch (ctype) {
		case NUMERIC:
			double d = cell.getNumericCellValue();
			int dd = new Double(d).intValue();
			if (d == new Double(dd).doubleValue()) {
				str = String.valueOf(dd);
			} else {
				str = String.valueOf(d);
			}
			break;
		case STRING:
			str = cell.getStringCellValue();
			break;
		case BOOLEAN:
			boolean bol = cell.getBooleanCellValue();
			str = String.valueOf(bol);
			break;
		default:
			break;
		}
		return str;
	}

	/**
	 * 读取String类型数据
	 * 
	 * @param value
	 *            数据
	 * @return [可能为空]
	 */
	public static String readString(String value) {
		return value;
	}

	/**
	 * 读取Boolean类型数据
	 * 
	 * @param cell
	 *            单元格
	 * @return [默认为False]
	 */
	public static Boolean readBoolean(Cell cell) {
		if (cell == null) {
			return Boolean.FALSE;
		}
		CellType ctype = cell.getCellTypeEnum();
		Boolean bol = null;
		switch (ctype) {
		case NUMERIC:
			double d = cell.getNumericCellValue();
			if (d == 1) {
				bol = true;
			} else if (d == 0) {
				bol = false;
			}
			break;
		case STRING:
			String d1 = cell.getStringCellValue();
			if ("true".equalsIgnoreCase(d1) || "1".equalsIgnoreCase(d1)) {
				bol = true;
			} else if ("false".equalsIgnoreCase(d1) || "0".equalsIgnoreCase(d1)) {
				bol = false;
			}
			break;
		case BOOLEAN:
			bol = cell.getBooleanCellValue();
			break;
		default:
			break;
		}
		return bol;
	}

	/**
	 * 读取Boolean类型数据
	 * 
	 * @param value
	 *            数据
	 * @return [默认为False]
	 */
	public static Boolean readBoolean(String value) {
		if (value == null) {
			return Boolean.FALSE;
		}
		if (value.equalsIgnoreCase("true")) {
			return true;
		} else if (value.equalsIgnoreCase("false")) {
			return true;
		} else {
			return value.charAt(0) == '0' ? false : true;
		}
	}

	/**
	 * 读取数据
	 * 
	 * @param fieldClass
	 *            读取的数据类型
	 * @param formatIndex
	 *            Excel数据类型
	 * @param value
	 *            数据
	 * @param express
	 *            数据格式
	 * @return 数据(读取日期时：如果是String写入，则根据此表达式进行格式化读取；读取Double、BigDecimal时，是保留的小数点后数字个数；)
	 */
	@SuppressWarnings("unchecked")
	public static <T> T read(Class<T> fieldClass, int formatIndex, String value, String express) {
		if (value == null) {
			return null;
		}
		Object obj = null;
		if (fieldClass.isAssignableFrom(int.class) || fieldClass.isAssignableFrom(Integer.class)) {
			obj = readInteger(value);
		} else if (fieldClass.isAssignableFrom(Double.class) || fieldClass.isAssignableFrom(double.class)) {
			obj = readDouble(value, ExcelUtils.parserExpressToInt(express));
		} else if (fieldClass.isAssignableFrom(Long.class) || fieldClass.isAssignableFrom(long.class)) {
			obj = readLong(value);
		} else if (fieldClass.isAssignableFrom(BigDecimal.class)) {
			obj = readBigDecimal(value, ExcelUtils.parserExpressToInt(express));
		} else if (fieldClass.isAssignableFrom(Date.class)) {
			obj = readDate(formatIndex, value, express);
		} else if (fieldClass.isAssignableFrom(Boolean.class) || fieldClass.isAssignableFrom(boolean.class)) {
			obj = readBoolean(value);
		} else if (fieldClass.isAssignableFrom(String.class)) {
			obj = readString(value);
		}
		return (T) obj;
	}

	/**
	 * 读取数据
	 * 
	 * @param fieldClass
	 *            读取的数据类型
	 * @param cell
	 *            单元格
	 * @param express
	 *            数据格式
	 * @return 数据(读取日期时：如果是String写入，则根据此表达式进行格式化读取；读取Double、BigDecimal时，是保留的小数点后数字个数；)
	 */
	@SuppressWarnings("unchecked")
	public static <T> T read(Class<T> fieldClass, Cell cell, String express) {
		if (cell == null) {
			return null;
		}
		Object obj = null;
		if (fieldClass.isAssignableFrom(int.class) || fieldClass.isAssignableFrom(Integer.class)) {
			obj = readInteger(cell);
		} else if (fieldClass.isAssignableFrom(Double.class) || fieldClass.isAssignableFrom(double.class)) {
			obj = readDouble(cell, ExcelUtils.parserExpressToInt(express));
		} else if (fieldClass.isAssignableFrom(Long.class) || fieldClass.isAssignableFrom(long.class)) {
			obj = readLong(cell);
		} else if (fieldClass.isAssignableFrom(BigDecimal.class)) {
			obj = readBigDecimal(cell, ExcelUtils.parserExpressToInt(express));
		} else if (fieldClass.isAssignableFrom(Date.class)) {
			obj = readDate(cell, express);
		} else if (fieldClass.isAssignableFrom(Boolean.class) || fieldClass.isAssignableFrom(boolean.class)) {
			obj = readBoolean(cell);
		} else if (fieldClass.isAssignableFrom(String.class)) {
			obj = readString(cell);
		}
		return (T) obj;
	}

}
