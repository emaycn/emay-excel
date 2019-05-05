package cn.emay.excel.read;

import java.math.BigDecimal;
import java.text.SimpleDateFormat;
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
public class ExcelReadHelper {

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
			date = parseDate(cell.getStringCellValue(), express);
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
				date = parseDate(value, express);
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
	 * 把字符串转成日期
	 * 
	 * @param dateStr
	 * @param format
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

}
