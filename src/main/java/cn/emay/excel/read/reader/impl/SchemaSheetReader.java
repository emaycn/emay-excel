package cn.emay.excel.read.reader.impl;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;

import cn.emay.excel.common.schema.base.ColumnSchema;
import cn.emay.excel.common.schema.base.SheetSchema;
import cn.emay.excel.read.ExcelReadHelper;
import cn.emay.excel.read.handler.SchemaSheetDataHandler;
import cn.emay.excel.read.handler.SheetDataHandler;
import cn.emay.excel.read.reader.SheetReader;

/**
 * 定义方式读取
 * 
 * @author Frank
 *
 * @param <D>
 *            数据
 */
public class SchemaSheetReader<D> implements SheetReader {

	/**
	 * 定义
	 */
	private SheetSchema<D> schema;
	/**
	 * 数据处理器
	 */
	private SheetDataHandler<D> dataReader;
	/**
	 * 是否按照列序号读取
	 */
	private boolean readByIndex;

	/**
	 * 当前sheet页的编号
	 */
	private int curSheetIndex;
	/**
	 * 当前sheet页的名字
	 */
	private String curSheetName;
	/**
	 * 当前数据
	 */
	private D curData;

	/**
	 * 所有的字段集合
	 */
	private Map<String, Field> fields = new HashMap<>();
	/**
	 * 列对应的标题
	 */
	private Map<Integer, String> colTitles = new HashMap<>();

	/**
	 * 
	 * @param schemaSheetDataHandler
	 *            基于定义的数据处理器
	 */
	public SchemaSheetReader(SchemaSheetDataHandler<D> schemaSheetDataHandler) {
		this(schemaSheetDataHandler.getSheetSchema(), schemaSheetDataHandler);
	}

	/**
	 * 
	 * @param schema
	 *            读取定义
	 * @param dataReader
	 *            数据处理器
	 */
	public SchemaSheetReader(SheetSchema<D> schema, SheetDataHandler<D> dataReader) {
		if (schema == null) {
			throw new IllegalArgumentException("schema is null");
		}
		schema.check();
		if (dataReader == null) {
			throw new IllegalArgumentException("dataReader is null");
		}
		this.schema = schema;
		this.dataReader = dataReader;
		this.readByIndex = schema.getSheetSchemaParams().readByIndex();
		Set<String> titles = new HashSet<>();
		Set<Integer> indexs = new HashSet<>();
		Field[] fieldArray = schema.getDataClass().getDeclaredFields();
		for (Field field : fieldArray) {
			field.setAccessible(true);
			ColumnSchema csma = schema.getExcelColumnByFieldName(field.getName());
			if (csma == null) {
				continue;
			}
			if (indexs.contains(csma.getIndex())) {
				throw new IllegalArgumentException(" has same columnIndex filed.");
			}
			indexs.add(csma.getIndex());
			if (titles.contains(csma.getTitle())) {
				throw new IllegalArgumentException(" has same title filed.");
			}
			titles.add(csma.getTitle());
			fields.put(readByIndex ? String.valueOf(csma.getIndex()) : csma.getTitle(), field);
		}
		if (fields.size() == 0) {
			throw new IllegalArgumentException(" has no filed to read");
		}
	}

	@Override
	public int getStartReadRowIndex() {
		return 0;
	}

	@Override
	public int getEndReadRowIndex() {
		return this.schema.getSheetSchemaParams().getReadDataEndRowIndex();
	}

	@Override
	public void begin(int sheetIndex, String sheetName) {
		this.curSheetIndex = sheetIndex;
		this.curSheetName = sheetName;
	}

	@Override
	public void beginRow(int rowIndex) {
		if (rowIndex < this.schema.getSheetSchemaParams().getReadDataStartRowIndex()) {
			return;
		}
		curData = schema.newData();
	}

	@Override
	public void handleXlsCell(int rowIndex, int columnIndex, Cell cell) {
		if (readByIndex == false && rowIndex == this.schema.getSheetSchemaParams().getReadTitleRowIndex()) {
			String title = ExcelReadHelper.readString(cell);
			colTitles.put(columnIndex, title == null ? "" : title);
		}
		if (rowIndex < this.schema.getSheetSchemaParams().getReadDataStartRowIndex()) {
			return;
		}
		if (cell == null) {
			return;
		}
		Field field = fields.get(readByIndex ? String.valueOf(columnIndex) : colTitles.get(columnIndex));
		if (field == null) {
			return;
		}
		ColumnSchema columnSchema = schema.getExcelColumnByFieldName(field.getName());
		if (columnSchema == null) {
			return;
		}
		Object obj = null;
		try {
			if (field.getType().isAssignableFrom(int.class) || field.getType().isAssignableFrom(Integer.class)) {
				obj = ExcelReadHelper.readInteger(cell);
			} else if (field.getType().isAssignableFrom(Double.class) || field.getType().isAssignableFrom(double.class)) {
				obj = ExcelReadHelper.readDouble(cell, columnSchema.getExpressInt());
			} else if (field.getType().isAssignableFrom(Long.class) || field.getType().isAssignableFrom(long.class)) {
				obj = ExcelReadHelper.readLong(cell);
			} else if (field.getType().isAssignableFrom(BigDecimal.class)) {
				obj = ExcelReadHelper.readBigDecimal(cell, columnSchema.getExpressInt());
			} else if (field.getType().isAssignableFrom(Date.class)) {
				obj = ExcelReadHelper.readDate(cell, columnSchema.getExpress());
			} else if (field.getType().isAssignableFrom(Boolean.class) || field.getType().isAssignableFrom(boolean.class)) {
				obj = ExcelReadHelper.readBoolean(cell);
			} else if (field.getType().isAssignableFrom(String.class)) {
				obj = ExcelReadHelper.readString(cell);
			}
			if (obj != null) {
				field.set(curData, obj);
			}
		} catch (IllegalArgumentException e) {
			throw new IllegalArgumentException(
					"sheet(" + curSheetName + "):[" + curSheetIndex + "]-row[" + rowIndex + "]-column[" + columnIndex + "] read[" + obj + "] and set[" + field.getName() + "] error", e);
		} catch (IllegalAccessException e) {
			throw new IllegalArgumentException(
					"sheet(" + curSheetName + "):[" + curSheetIndex + "]-row[" + rowIndex + "]-column[" + columnIndex + "] read[" + obj + "] and set[" + field.getName() + "] error", e);
		}
	}

	@Override
	public void handleXlsxCell(int rowIndex, int columnIndex, int formatIndex, String value) {
		if (readByIndex == false && rowIndex == this.schema.getSheetSchemaParams().getReadTitleRowIndex()) {
			String title = ExcelReadHelper.readString(value);
			colTitles.put(columnIndex, title == null ? "" : title);
		}
		if (rowIndex < this.schema.getSheetSchemaParams().getReadDataStartRowIndex()) {
			return;
		}
		if (value == null) {
			return;
		}
		Field field = fields.get(readByIndex ? String.valueOf(columnIndex) : colTitles.get(columnIndex));
		if (field == null) {
			return;
		}
		ColumnSchema columnSchema = schema.getExcelColumnByFieldName(field.getName());
		if (columnSchema == null) {
			return;
		}
		Object obj = null;
		try {
			if (field.getType().isAssignableFrom(int.class) || field.getType().isAssignableFrom(Integer.class)) {
				obj = ExcelReadHelper.readInteger(value);
			} else if (field.getType().isAssignableFrom(Double.class) || field.getType().isAssignableFrom(double.class)) {
				obj = ExcelReadHelper.readDouble(value, columnSchema.getExpressInt());
			} else if (field.getType().isAssignableFrom(Long.class) || field.getType().isAssignableFrom(long.class)) {
				obj = ExcelReadHelper.readLong(value);
			} else if (field.getType().isAssignableFrom(BigDecimal.class)) {
				obj = ExcelReadHelper.readBigDecimal(value, columnSchema.getExpressInt());
			} else if (field.getType().isAssignableFrom(Date.class)) {
				obj = ExcelReadHelper.readDate(formatIndex, value, columnSchema.getExpress());
			} else if (field.getType().isAssignableFrom(Boolean.class) || field.getType().isAssignableFrom(boolean.class)) {
				obj = ExcelReadHelper.readBoolean(value);
			} else if (field.getType().isAssignableFrom(String.class)) {
				obj = ExcelReadHelper.readString(value);
			}
			if (obj != null) {
				field.set(curData, obj);
			}
		} catch (IllegalArgumentException e) {
			throw new IllegalArgumentException(
					"sheet(" + curSheetName + "):[" + curSheetIndex + "]-row[" + rowIndex + "]-column[" + columnIndex + "] read[" + obj + "] and set[" + field.getName() + "] error", e);
		} catch (IllegalAccessException e) {
			throw new IllegalArgumentException(
					"sheet(" + curSheetName + "):[" + curSheetIndex + "]-row[" + rowIndex + "]-column[" + columnIndex + "] read[" + obj + "] and set[" + field.getName() + "] error", e);
		}
	}

	@Override
	public void endRow(int rowIndex) {
		if (curData != null) {
			dataReader.handle(rowIndex, curData);
		}
	}

	@Override
	public void end(int sheetIndex, String sheetName) {
		curData = null;
	}
}