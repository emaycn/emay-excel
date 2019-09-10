package cn.emay.excel.write.writer.impl;

import java.lang.reflect.Field;
import java.util.Arrays;
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

import cn.emay.excel.common.schema.base.ColumnSchema;
import cn.emay.excel.common.schema.base.SheetSchema;
import cn.emay.excel.common.schema.base.SheetWriteSchemaParams;
import cn.emay.excel.utils.ExcelWriteUtils;
import cn.emay.excel.write.data.SheetDataGetter;
import cn.emay.excel.write.writer.SheetWriter;

/**
 * 
 * 基于Schema的写入处理器
 * 
 * @author Frank
 *
 * @param <D>
 */
public class SchemaSheetWriter<D> implements SheetWriter {

	/**
	 * 默认颜色
	 */
	private static int[] DEFAULT_RGB_COLOR = { 255, 255, 255 };

	/**
	 * 每一列的最大宽度
	 */
	private Map<Integer, Integer> maxWidth = new HashMap<>();
	/**
	 * 是否内容需要颜色
	 */
	private boolean isContentNeedColor;
	/**
	 * 是否标题需要颜色
	 */
	private boolean isTitleNeedColor;

	/**
	 * 读定义参数集
	 */
	private SheetWriteSchemaParams writeSchemaParams;
	/**
	 * 列定义集合
	 */
	private Map<Integer, ColumnSchema> schemaMap = new HashMap<>();
	/**
	 * 字段集合
	 */
	private Map<Integer, Field> fieldMap = new HashMap<>();
	/**
	 * 最大的列序号
	 */
	private int maxColumnIndex = 0;
	/**
	 * 当前处理的Sheet页序号
	 */
	private int sheetIndex;
	/**
	 * 数据写入处理器
	 */
	private SheetDataGetter<D> writeData;
	/**
	 * 当前数据
	 */
	private D curr;
	/**
	 * sheet对象
	 */
	private Sheet sheet;

	/**
	 * 
	 * @param writeData
	 *            数据写入处理器
	 */
	public SchemaSheetWriter(SheetSchema schema, SheetDataGetter<D> writeData) {
		if (schema == null) {
			throw new IllegalArgumentException("schema is null");
		}
		schema.checkWrite();
		if (writeData == null) {
			throw new IllegalArgumentException("writeData is null");
		}
		this.writeData = writeData;
		// this.schema = schema;
		this.writeSchemaParams = schema.getSheetWriteSchemaParams();
		this.isContentNeedColor = !Arrays.equals(writeSchemaParams.getContentRgbColor(), DEFAULT_RGB_COLOR);
		this.isTitleNeedColor = !Arrays.equals(writeSchemaParams.getTitleRgbColor(), DEFAULT_RGB_COLOR);
		Set<Integer> columnIndexs = new HashSet<>();
		Field[] fields = writeData.getDataClass().getDeclaredFields();
		for (Field field : fields) {
			field.setAccessible(true);
			ColumnSchema csma = schema.getExcelColumnByFieldName(field.getName());
			if (csma == null) {
				continue;
			}
			if (columnIndexs.contains(csma.getIndex())) {
				throw new IllegalArgumentException(writeData.getDataClass().getName() + " has same columnIndex[" + csma.getIndex() + "] filed.");
			}
			columnIndexs.add(csma.getIndex());
			schemaMap.put(csma.getIndex(), csma);
			fieldMap.put(csma.getIndex(), field);
			maxColumnIndex = maxColumnIndex > csma.getIndex() ? maxColumnIndex : csma.getIndex();
		}
		if (fieldMap.size() == 0) {
			throw new IllegalArgumentException("dataClass[" + writeData.getDataClass().getName() + "] is not has ExcelColumn filed  ");
		}
	}

	@Override
	public String getSheetName() {
		return writeSchemaParams.getWriteSheetName();
	}

	@Override
	public boolean hasRow(int rowIndex) {
		if (writeSchemaParams.isWriteTile()) {
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
		if (writeSchemaParams.isWriteTile()) {
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
		ColumnSchema columnSchema = schemaMap.get(columnIndex);
		if (columnSchema == null) {
			return;
		}
		int length = 0;
		if (rowIndex == 0 && writeSchemaParams.isWriteTile()) {
			String title = "".equals(columnSchema.getTitle().trim()) ? field.getName() : columnSchema.getTitle();
			ExcelWriteUtils.writeString(cell, title);
			length = title.getBytes().length;
		} else {
			if (curr == null) {
				return;
			}
			try {
				Object obj = field.get(curr);
				ExcelWriteUtils.write(cell, obj, columnSchema.getExpress());
				if (writeSchemaParams.isAutoWidth()) {
					if (!boolean.class.isAssignableFrom(field.getType()) && !Boolean.class.isAssignableFrom(field.getType())) {
						length = String.valueOf(obj).getBytes().length;
					} else {
						length = 6;
					}
				}
			} catch (IllegalArgumentException e) {
				throw new IllegalArgumentException("sheet(" + writeSchemaParams.getWriteSheetName() + ")[" + sheetIndex + "]-row[" + rowIndex + "]-column[" + columnIndex + "] get value from [" + field.getName() + "] and write error", e);
			} catch (IllegalAccessException e) {
				throw new IllegalArgumentException("sheet(" + writeSchemaParams.getWriteSheetName() + ")[" + sheetIndex + "]-row[" + rowIndex + "]-column[" + columnIndex + "] get value from [" + field.getName() + "] and write error", e);
			}
		}
		if (writeSchemaParams.isAutoWidth()) {
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
		if (writeSchemaParams.isAutoWidth()) {
			for (Integer columnIndex : maxWidth.keySet()) {
				Integer width = maxWidth.get(columnIndex);
				if (width != null) {
					sheet.setColumnWidth(columnIndex, width * 125 / 100);
				}
			}
		}
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
			if (writeSchemaParams.isAutoWidth()) {
				sheet.autoSizeColumn(columnIndex);
			}
			if (HSSFWorkbook.class.isAssignableFrom(cell.getSheet().getWorkbook().getClass())) {
				HSSFPalette palette = ((HSSFWorkbook) cell.getSheet().getWorkbook()).getCustomPalette();
				if (isTitleNeedColor) {
					palette.setColorAtIndex(HSSFColorPredefined.GREY_25_PERCENT.getIndex(), (byte) writeSchemaParams.getTitleRgbColor()[0], (byte) writeSchemaParams.getTitleRgbColor()[1], (byte) writeSchemaParams.getTitleRgbColor()[2]);
				}
				if (isContentNeedColor) {
					palette.setColorAtIndex(HSSFColorPredefined.GREY_40_PERCENT.getIndex(), (byte) writeSchemaParams.getContentRgbColor()[0], (byte) writeSchemaParams.getContentRgbColor()[1],
							(byte) writeSchemaParams.getContentRgbColor()[2]);
				}
			}
		}
		if (rowIndex == 0 && writeSchemaParams.isWriteTile()) {
			Font font = cell.getSheet().getWorkbook().createFont();
			font.setBold(true);
			style.setFont(font);
			style.setAlignment(HorizontalAlignment.CENTER);
			if (isTitleNeedColor) {
				if (HSSFWorkbook.class.isAssignableFrom(cell.getSheet().getWorkbook().getClass())) {
					style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
				} else {
					((XSSFCellStyle) style).setFillForegroundColor(new XSSFColor(new java.awt.Color(writeSchemaParams.getTitleRgbColor()[0], writeSchemaParams.getTitleRgbColor()[1], writeSchemaParams.getTitleRgbColor()[2])));
				}
				style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			}
		} else if (isContentNeedColor) {
			if (HSSFWorkbook.class.isAssignableFrom(cell.getSheet().getWorkbook().getClass())) {
				style.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
			} else {
				((XSSFCellStyle) style).setFillForegroundColor(new XSSFColor(new java.awt.Color(writeSchemaParams.getContentRgbColor()[0], writeSchemaParams.getContentRgbColor()[1], writeSchemaParams.getContentRgbColor()[2])));
			}
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		}

		if (writeSchemaParams.isNeedBorder()) {
			style.setBorderLeft(BorderStyle.THIN);
			style.setBorderTop(BorderStyle.THIN);
			style.setBorderBottom(BorderStyle.THIN);
			style.setBorderRight(BorderStyle.THIN);
		}
		if (writeSchemaParams.isAutoWrap()) {
			style.setWrapText(true);
		}
		cell.setCellStyle(style);
	}

}
