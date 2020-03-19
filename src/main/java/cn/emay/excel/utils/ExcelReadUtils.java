package cn.emay.excel.utils;

import java.math.BigDecimal;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.ExcelNumberFormat;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 读工具类
 * 
 * @author Frank
 *
 */
public class ExcelReadUtils {

	/**
	 * WORKBOOK模式浮点数格式
	 */
	private static HSSFDataFormatter HDF = new HSSFDataFormatter();

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
		case FORMULA:
			FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
			CellValue value = evaluator.evaluate(cell);
			if (value == null) {
				break;
			}
			switch (value.getCellTypeEnum()) {
			case NUMERIC:
				Workbook wb = cell.getSheet().getWorkbook();
				boolean date1904 = false;
				if (wb.getClass().getName().equals(SXSSFWorkbook.class.getName())) {
					date1904 = ((SXSSFWorkbook) wb).getXSSFWorkbook().isDate1904();
				} else if (wb.getClass().getName().equals(XSSFWorkbook.class.getName())) {
					date1904 = ((XSSFWorkbook) wb).isDate1904();
				} else if (wb.getClass().getName().equals(HSSFWorkbook.class.getName())) {
					date1904 = ((HSSFWorkbook) wb).getInternalWorkbook().isUsing1904DateWindowing();
				}
				date = DateUtil.getJavaDate(value.getNumberValue(), date1904);
				break;
			case STRING:
				date = ExcelUtils.parseDate(value.getStringValue(), express);
				break;
			default:
				break;
			}
			break;
		default:
			break;
		}
		return date;
	}

	/**
	 * 读取日期类型数据
	 * 
	 * @param value
	 *            数据
	 * @param format
	 *            日期格式
	 * @return [可能为空]
	 */
	public static Date readDate(String value, String express) {
		if (value == null) {
			return null;
		}
		Date date = null;
		try {
			date = ExcelUtils.parseDate(value, express);
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
		case FORMULA:
			FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
			CellValue value = evaluator.evaluate(cell);
			if (value == null) {
				break;
			}
			switch (value.getCellTypeEnum()) {
			case NUMERIC:
				d2 = new BigDecimal(value.getNumberValue());
				break;
			case STRING:
				try {
					d2 = new BigDecimal(value.getStringValue());
				} catch (Exception e) {
				}
				break;
			case BOOLEAN:
				d2 = new BigDecimal(value.getBooleanValue() ? 1d : 0d);
				break;
			default:
				break;
			}
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
			str = HDF.formatCellValue(cell);
			break;
		case STRING:
			str = cell.getStringCellValue();
			break;
		case BOOLEAN:
			boolean bol = cell.getBooleanCellValue();
			str = String.valueOf(bol);
			break;
		case FORMULA:
			FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
			CellValue value = evaluator.evaluate(cell);
			if (value == null) {
				break;
			}
			switch (value.getCellTypeEnum()) {
			case NUMERIC:
				ExcelNumberFormat numFmt = ExcelNumberFormat.from(cell, null);
				str = HDF.formatRawCellContents(value.getNumberValue(), numFmt.getIdx(), numFmt.getFormat());
				break;
			case STRING:
				str = value.getStringValue();
				break;
			case BOOLEAN:
				boolean bol1 = value.getBooleanValue();
				str = String.valueOf(bol1);
				break;
			default:
				break;
			}
			break;
		default:
			break;
		}
		return str;
	}

	/**
	 * 读取String类型数据<br/>
	 * 浮点数精度问题通过Double.value()解决,会损失部分性能
	 * 
	 * @param value
	 *            数据
	 * @return [可能为空]
	 */
	public static String readString(String value) {
		if(value == null || "".equals(value)) {
			return value;
		}
		if(value.startsWith("0")) {
			return value;
		}
		// 先进行整数解析，如果匹配上了，直接返回
		try {
			Long lon = Long.valueOf(value);
			return lon.toString();
		} catch (Exception e) {
		}
		// 再进行小数解析，如果匹配上了，直接返回
		try {
			Double dou = Double.valueOf(value);
			return dou.toString();
		} catch (Exception e) {
		}
		// 如果没有解析成数字，放回原值
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
		case FORMULA:
			FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
			CellValue value = evaluator.evaluate(cell);
			if (value == null) {
				break;
			}
			switch (value.getCellTypeEnum()) {
			case NUMERIC:
				double dv = value.getNumberValue();
				if (dv == 1) {
					bol = true;
				} else if (dv == 0) {
					bol = false;
				}
				break;
			case STRING:
				String d12 = value.getStringValue();
				if ("true".equalsIgnoreCase(d12) || "1".equalsIgnoreCase(d12)) {
					bol = true;
				} else if ("false".equalsIgnoreCase(d12) || "0".equalsIgnoreCase(d12)) {
					bol = false;
				}
				break;
			case BOOLEAN:
				bol = value.getBooleanValue();
				break;
			default:
				break;
			}
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
			return false;
		} else {
			return value.charAt(0) == '0' ? false : true;
		}
	}

	/**
	 * 读取数据
	 * 
	 * @param fieldClass
	 *            读取的数据类型
	 * @param value
	 *            数据
	 * @param express
	 *            数据格式
	 * @return 数据(读取日期时：如果是String写入，则根据此表达式进行格式化读取；读取Double、BigDecimal时，是保留的小数点后数字个数；)
	 */
	@SuppressWarnings("unchecked")
	public static <T> T read(Class<T> fieldClass, String value, String express) {
		if (value == null) {
			return null;
		}
		Object obj = null;
		if (int.class.isAssignableFrom(fieldClass) || Integer.class.isAssignableFrom(fieldClass)) {
			obj = readInteger(value);
		} else if (Double.class.isAssignableFrom(fieldClass) || double.class.isAssignableFrom(fieldClass)) {
			obj = readDouble(value, ExcelUtils.parserExpressToInt(express));
		} else if (Long.class.isAssignableFrom(fieldClass) || long.class.isAssignableFrom(fieldClass)) {
			obj = readLong(value);
		} else if (BigDecimal.class.isAssignableFrom(fieldClass)) {
			obj = readBigDecimal(value, ExcelUtils.parserExpressToInt(express));
		} else if (Date.class.isAssignableFrom(fieldClass)) {
			obj = readDate(value, express);
		} else if (Boolean.class.isAssignableFrom(fieldClass) || boolean.class.isAssignableFrom(fieldClass)) {
			obj = readBoolean(value);
		} else if (String.class.isAssignableFrom(fieldClass)) {
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
		if (int.class.isAssignableFrom(fieldClass) || Integer.class.isAssignableFrom(fieldClass)) {
			obj = readInteger(cell);
		} else if (Double.class.isAssignableFrom(fieldClass) || double.class.isAssignableFrom(fieldClass)) {
			obj = readDouble(cell, ExcelUtils.parserExpressToInt(express));
		} else if (Long.class.isAssignableFrom(fieldClass) || long.class.isAssignableFrom(fieldClass)) {
			obj = readLong(cell);
		} else if (BigDecimal.class.isAssignableFrom(fieldClass)) {
			obj = readBigDecimal(cell, ExcelUtils.parserExpressToInt(express));
		} else if (Date.class.isAssignableFrom(fieldClass)) {
			obj = readDate(cell, express);
		} else if (Boolean.class.isAssignableFrom(fieldClass) || boolean.class.isAssignableFrom(fieldClass)) {
			obj = readBoolean(cell);
		} else if (String.class.isAssignableFrom(fieldClass)) {
			obj = readString(cell);
		}
		return (T) obj;
	}

}
