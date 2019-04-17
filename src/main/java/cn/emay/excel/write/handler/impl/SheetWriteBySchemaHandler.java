package cn.emay.excel.write.handler.impl;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

import cn.emay.excel.common.ExcelColumn;
import cn.emay.excel.common.ExcelSheet;
import cn.emay.excel.write.ExcelWriterHelper;
import cn.emay.excel.write.handler.SheetWriteHandler;

/**
 * 
 * 基于Schema的写入处理器
 * 
 * @author Frank
 *
 * @param <D>
 */
public class SheetWriteBySchemaHandler<D extends Object> implements SheetWriteHandler {

	/**
	 * 默认颜色
	 */
	private static int[] DEFAULT_RGB_COLOR = { 255, 255, 255 };

	/**
	 * Excel定义
	 */
	private ExcelSheet schema;
	/**
	 * Excel名字
	 */
	private String name;
	/**
	 * 当前处理的Sheet页序号
	 */
	private int sheetIndex;
	/**
	 * 最大的列序号
	 */
	private int maxColumnIndex;
	/**
	 * 数据写入处理器
	 */
	private DataWriter<D> writeData;
	/**
	 * 列定义集合
	 */
	private Map<Integer, ExcelColumn> schemaMap;
	/**
	 * 字段集合
	 */
	private Map<Integer, Field> fieldMap;
	/**
	 * 当前数据
	 */
	private D curr;
	/**
	 * 每一列的最大宽度
	 */
	private Map<Integer, Integer> maxWidth = new HashMap<>();
	/**
	 * sheet对象
	 */
	private Sheet sheet;
	/**
	 * 样式对象
	 */
	private HSSFPalette palette;
	/**
	 * 是否内容需要颜色
	 */
	private boolean isContentNeedColor;
	/**
	 * 是否标题需要颜色
	 */
	private boolean isTitleNeedColor;

	/**
	 * 
	 * @param writeData
	 *            数据写入处理器
	 */
	public SheetWriteBySchemaHandler(DataWriter<D> writeData) {
		if (writeData.getShemaClass() == null) {
			throw new IllegalArgumentException("schemaClass is null");
		}
		if (!writeData.getShemaClass().isAnnotationPresent(ExcelSheet.class)) {
			throw new IllegalArgumentException("schemaClass[" + writeData.getShemaClass().getName() + "] is not has Annotation : " + ExcelSheet.class.getName());
		}
		ExcelSheet schema = writeData.getShemaClass().getAnnotation(ExcelSheet.class);
		Field[] fields = writeData.getShemaClass().getDeclaredFields();
		if (fields == null || fields.length == 0) {
			throw new IllegalArgumentException("schemaClass[" + writeData.getShemaClass().getName() + "] is not has  filed  ");
		}
		Map<Integer, ExcelColumn> schemaMap = new HashMap<>();
		Map<Integer, Field> fieldMap = new HashMap<>();
		Set<Integer> columnIndexs = new HashSet<>();
		int maxColumnIndex = 0;
		for (Field field : fields) {
			if (!field.isAnnotationPresent(ExcelColumn.class)) {
				continue;
			}
			ExcelColumn csma = field.getAnnotation(ExcelColumn.class);
			if (columnIndexs.contains(csma.columnIndex())) {
				throw new IllegalArgumentException(writeData.getShemaClass().getName() + " has same columnIndex filed.");
			}
			columnIndexs.add(csma.columnIndex());
			schemaMap.put(csma.columnIndex(), csma);
			fieldMap.put(csma.columnIndex(), field);
			field.setAccessible(true);
			maxColumnIndex = maxColumnIndex > csma.columnIndex() ? maxColumnIndex : csma.columnIndex();
		}
		if (schemaMap.size() == 0) {
			throw new IllegalArgumentException("schemaClass[" + writeData.getShemaClass().getName() + "] is not has ExcelColumnSchema filed  ");
		}
		this.schema = schema;
		this.maxColumnIndex = maxColumnIndex;
		this.name = schema.writeSheetName();
		this.schemaMap = schemaMap;
		this.fieldMap = fieldMap;
		this.writeData = writeData;
		isContentNeedColor = !Arrays.equals(schema.contentRgbColor(), DEFAULT_RGB_COLOR);
		isTitleNeedColor = !Arrays.equals(schema.titleRgbColor(), DEFAULT_RGB_COLOR);
	}

	@Override
	public String getSheetName() {
		return name;
	}

	@Override
	public boolean hasRow(int rowIndex) {
		if (schema.isWriteTile()) {
			if (rowIndex == 0) {
				return true;
			} else {
				return writeData.hasData(rowIndex - 1);
			}
		} else {
			return writeData.hasData(rowIndex);
		}

	}

	@Override
	public int getMaxColumnIndex() {
		return maxColumnIndex;
	}

	@Override
	public void begin(int sheetIndex) {
		this.sheetIndex = sheetIndex;
	}

	@Override
	public void beginRow(int rowIndex) {
		if (schema.isWriteTile()) {
			if (rowIndex != 0) {
				curr = writeData.getData(rowIndex - 1);
			}
		} else {
			curr = writeData.getData(rowIndex);
		}
	}

	@Override
	public void writeCell(Cell cell, int rowIndex, int columnIndex) {
		writeStyle(cell, rowIndex, columnIndex);
		Field field = fieldMap.get(columnIndex);
		if (field == null) {
			return;
		}
		ExcelColumn columnSchema = schemaMap.get(columnIndex);
		if (columnSchema == null) {
			return;
		}
		int length = 0;
		if (rowIndex == 0 && schema.isWriteTile()) {
			String title = "".equals(columnSchema.title().trim()) ? field.getName() : columnSchema.title();
			ExcelWriterHelper.writeString(cell, title);
			length = title.getBytes().length;
		} else {
			if (curr == null) {
				return;
			}
			try {
				Object obj = field.get(curr);
				if (obj != null) {
					if (field.getType().isAssignableFrom(int.class)) {
						ExcelWriterHelper.writeInt(cell, (int) obj);
						length = getLength(obj);
					} else if (field.getType().isAssignableFrom(Integer.class)) {
						ExcelWriterHelper.writeInt(cell, (Integer) obj);
						length = getLength(obj);
					} else if (field.getType().isAssignableFrom(double.class)) {
						ExcelWriterHelper.writeDouble(cell, (double) obj, parseNumberExpress(columnSchema.express()));
						length = getLength(obj);
					} else if (field.getType().isAssignableFrom(Double.class)) {
						ExcelWriterHelper.writeDouble(cell, (Double) obj, parseNumberExpress(columnSchema.express()));
						length = getLength(obj);
					} else if (field.getType().isAssignableFrom(long.class)) {
						ExcelWriterHelper.writeLong(cell, (long) obj);
						length = getLength(obj);
					} else if (field.getType().isAssignableFrom(Long.class)) {
						ExcelWriterHelper.writeLong(cell, (Long) obj);
						length = getLength(obj);
					} else if (field.getType().isAssignableFrom(BigDecimal.class)) {
						ExcelWriterHelper.writeBigDecimal(cell, (BigDecimal) obj, parseNumberExpress(columnSchema.express()));
						length = getLength(((BigDecimal) obj).doubleValue());
					} else if (field.getType().isAssignableFrom(Date.class)) {
						ExcelWriterHelper.writeDate(cell, (Date) obj, columnSchema.express());
						length = getLength(columnSchema.express().trim().equals("") ? obj : columnSchema.express());
					} else if (field.getType().isAssignableFrom(boolean.class)) {
						ExcelWriterHelper.writeBoolean(cell, (boolean) obj);
						length = 6;
					} else if (field.getType().isAssignableFrom(Boolean.class)) {
						ExcelWriterHelper.writeBoolean(cell, (Boolean) obj);
						length = 6;
					} else if (field.getType().isAssignableFrom(String.class)) {
						ExcelWriterHelper.writeString(cell, (String) obj);
						length = getLength(obj);
					}
				}
			} catch (IllegalArgumentException e) {
				throw new IllegalArgumentException("sheet(" + name + ")[" + sheetIndex + "]-row[" + rowIndex + "]-column[" + columnIndex + "] get value from [" + field.getName() + "] and write error",
						e);
			} catch (IllegalAccessException e) {
				throw new IllegalArgumentException("sheet(" + name + ")[" + sheetIndex + "]-row[" + rowIndex + "]-column[" + columnIndex + "] get value from [" + field.getName() + "] and write error",
						e);
			}
		}
		if (schema.isAutoWidth()) {
			length *= 256;
			Integer maxlength = maxWidth.get(columnIndex);
			if (maxlength == null || maxlength.intValue() < length) {
				maxWidth.put(columnIndex, length);
			}
		}
	}

	@Override
	public void endRow(int rowIndex) {
		curr = null;
	}

	@Override
	public void end(int sheetIndex) {
		if (schema.isAutoWidth()) {
			for (Integer columnIndex : maxWidth.keySet()) {
				Integer width = maxWidth.get(columnIndex);
				if (width != null) {
					sheet.setColumnWidth(columnIndex, width * 125 / 100);
				}
			}
		}
	}

	/**
	 * 获取数据长度
	 * 
	 * @param obj
	 *            数据
	 * @return
	 */
	private int getLength(Object obj) {
		if (schema.isAutoWidth()) {
			return String.valueOf(obj).getBytes().length;
		}
		return 0;
	}

	/**
	 * 解析schema的express小数点后个数
	 * 
	 * @param express
	 *            表达式
	 * @return
	 */
	private int parseNumberExpress(String express) {
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
	 * 写入样式
	 * 
	 * @param cell
	 *            单元格
	 * @param rowIndex
	 *            行号
	 * @param columnIndex
	 *            列号
	 */
	private void writeStyle(Cell cell, int rowIndex, int columnIndex) {
		CellStyle style = cell.getCellStyle();
		if (rowIndex == 0) {
			sheet = cell.getSheet();
			if (schema.isAutoWidth()) {
				sheet.autoSizeColumn(columnIndex);
			}
			if (cell.getSheet().getWorkbook().getClass().isAssignableFrom(HSSFWorkbook.class)) {
				palette = ((HSSFWorkbook) cell.getSheet().getWorkbook()).getCustomPalette();
				if (isTitleNeedColor) {
					palette.setColorAtIndex(HSSFColorPredefined.GREY_25_PERCENT.getIndex(), (byte) schema.titleRgbColor()[0], (byte) schema.titleRgbColor()[1], (byte) schema.titleRgbColor()[2]);
				}
				if (isContentNeedColor) {
					palette.setColorAtIndex(HSSFColorPredefined.GREY_40_PERCENT.getIndex(), (byte) schema.contentRgbColor()[0], (byte) schema.contentRgbColor()[1], (byte) schema.contentRgbColor()[2]);
				}
			}
		}
		if (rowIndex == 0 && schema.isWriteTile()) {
			Font font = cell.getSheet().getWorkbook().createFont();
			font.setBold(true);
			style.setFont(font);
			style.setAlignment(HorizontalAlignment.CENTER);
			if (isTitleNeedColor) {
				if (cell.getSheet().getWorkbook().getClass().isAssignableFrom(HSSFWorkbook.class)) {
					style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
				} else {
					((XSSFCellStyle) style).setFillForegroundColor(new XSSFColor(new java.awt.Color(schema.titleRgbColor()[0], schema.titleRgbColor()[1], schema.titleRgbColor()[2])));
				}
				style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			}
		} else if (isContentNeedColor) {
			if (cell.getSheet().getWorkbook().getClass().isAssignableFrom(HSSFWorkbook.class)) {
				style.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
			} else {
				((XSSFCellStyle) style).setFillForegroundColor(new XSSFColor(new java.awt.Color(schema.contentRgbColor()[0], schema.contentRgbColor()[1], schema.contentRgbColor()[2])));
			}
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		}

		if (schema.isNeedBorder()) {
			style.setBorderLeft(BorderStyle.THIN);
			style.setBorderTop(BorderStyle.THIN);
			style.setBorderBottom(BorderStyle.THIN);
			style.setBorderRight(BorderStyle.THIN);
		}
		if (schema.isAutoWrap()) {
			style.setWrapText(true);
		}
		cell.setCellStyle(style);
	}

}
