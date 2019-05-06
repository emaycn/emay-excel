package cn.emay.excel.utils;

import java.math.BigDecimal;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;

/**
 * 写工具类
 * 
 * @author Frank
 *
 */
public class ExcelWriteUtils {

	/**
	 * 写入日期
	 * 
	 * @param cell
	 *            单元格
	 * @param date
	 *            日期
	 * @param format
	 *            日期格式
	 */
	public static void writeDate(Cell cell, Date date, String format) {
		if (cell == null) {
			return;
		}
		cell.setCellType(CellType.NUMERIC);
		if (date == null) {
			return;
		}
		double datenum = HSSFDateUtil.getExcelDate(date);
		cell.setCellValue(datenum);
		if (format != null) {
			CellStyle style = cell.getCellStyle();
			short df = cell.getSheet().getWorkbook().createDataFormat().getFormat(format);
			style.setDataFormat(df);
			cell.setCellStyle(style);
		}
	}

	/**
	 * 写入布尔
	 * 
	 * @param cell
	 *            单元格
	 * @param bool
	 *            布尔值
	 */
	public static void writeBoolean(Cell cell, boolean bool) {
		if (cell == null) {
			return;
		}
		cell.setCellType(CellType.BOOLEAN);
		cell.setCellValue(bool);
	}

	/**
	 * 写入浮点数
	 * 
	 * @param cell
	 *            单元格
	 * @param number
	 *            浮点数
	 */
	public static void writeDouble(Cell cell, double number) {
		writeDouble(cell, number, -1);
	}

	/**
	 * 写入浮点数
	 * 
	 * @param cell
	 *            单元格
	 * @param number
	 *            浮点数
	 * @param scale
	 *            保留小数点后位数。(<0则不改变原有值)
	 */
	public static void writeDouble(Cell cell, double number, int scale) {
		writeBigDecimal(cell, new BigDecimal(number), scale);
	}

	/**
	 * 写入浮点数
	 * 
	 * @param cell
	 *            单元格
	 * @param number
	 *            浮点数
	 */
	public static void writeBigDecimal(Cell cell, BigDecimal number) {
		writeBigDecimal(cell, number, -1);
	}

	/**
	 * 写入浮点数
	 * 
	 * @param cell
	 *            单元格
	 * @param number
	 *            浮点数
	 * @param scale
	 *            保留小数点后位数。(<0则不改变原有值)
	 */
	public static void writeBigDecimal(Cell cell, BigDecimal number, int scale) {
		if (cell == null) {
			return;
		}
		cell.setCellType(CellType.NUMERIC);
		if (number == null) {
			return;
		}
		if (scale >= 0) {
			cell.setCellValue(number.setScale(scale, BigDecimal.ROUND_HALF_UP).doubleValue());
		} else {
			cell.setCellValue(number.doubleValue());
		}
	}

	/**
	 * 写入长整型
	 * 
	 * @param cell
	 *            单元格
	 * @param number
	 *            长整型值
	 */
	public static void writeLong(Cell cell, long number) {
		if (cell == null) {
			return;
		}
		cell.setCellType(CellType.NUMERIC);
		cell.setCellValue(number);
		CellStyle style = cell.getCellStyle();
		// style.cloneStyleFrom(cell.getCellStyle());
		short df = cell.getSheet().getWorkbook().createDataFormat().getFormat("0");
		style.setDataFormat(df);
		cell.setCellStyle(style);
	}

	/**
	 * 写入整型
	 * 
	 * @param cell
	 *            单元格
	 * @param number
	 *            整型值
	 */
	public static void writeInt(Cell cell, int number) {
		if (cell == null) {
			return;
		}
		cell.setCellType(CellType.NUMERIC);
		cell.setCellValue(number);
		CellStyle style = cell.getCellStyle();
		short df = cell.getSheet().getWorkbook().createDataFormat().getFormat("0");
		style.setDataFormat(df);
		cell.setCellStyle(style);
	}

	/**
	 * 写入字符串
	 * 
	 * @param cell
	 *            单元格
	 * @param value
	 *            字符串
	 */
	public static void writeString(Cell cell, String value) {
		if (cell == null) {
			return;
		}
		cell.setCellType(CellType.STRING);
		if (value == null) {
			return;
		}
		cell.setCellValue(value);
	}

	/**
	 * 写入数据
	 * 
	 * @param cell
	 *            单元格
	 * @param data
	 *            数据
	 * @param format
	 *            格式
	 */
	public static void write(Cell cell, Object data, String format) {
		if (data == null) {
			return;
		}
		if (data.getClass().isAssignableFrom(int.class)) {
			ExcelWriteUtils.writeInt(cell, (int) data);
		} else if (data.getClass().isAssignableFrom(Integer.class)) {
			ExcelWriteUtils.writeInt(cell, (Integer) data);
		} else if (data.getClass().isAssignableFrom(double.class)) {
			ExcelWriteUtils.writeDouble(cell, (double) data, ExcelUtils.parserExpressToInt(format));
		} else if (data.getClass().isAssignableFrom(Double.class)) {
			ExcelWriteUtils.writeDouble(cell, (Double) data, ExcelUtils.parserExpressToInt(format));
		} else if (data.getClass().isAssignableFrom(long.class)) {
			ExcelWriteUtils.writeLong(cell, (long) data);
		} else if (data.getClass().isAssignableFrom(Long.class)) {
			ExcelWriteUtils.writeLong(cell, (Long) data);
		} else if (data.getClass().isAssignableFrom(BigDecimal.class)) {
			ExcelWriteUtils.writeBigDecimal(cell, (BigDecimal) data, ExcelUtils.parserExpressToInt(format));
		} else if (data.getClass().isAssignableFrom(Date.class)) {
			ExcelWriteUtils.writeDate(cell, (Date) data, format);
		} else if (data.getClass().isAssignableFrom(boolean.class)) {
			ExcelWriteUtils.writeBoolean(cell, (boolean) data);
		} else if (data.getClass().isAssignableFrom(Boolean.class)) {
			ExcelWriteUtils.writeBoolean(cell, (Boolean) data);
		} else if (data.getClass().isAssignableFrom(String.class)) {
			ExcelWriteUtils.writeString(cell, (String) data);
		}
	}

}
