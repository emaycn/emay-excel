package cn.emay.excel.read.handler.relize;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;

import cn.emay.excel.common.ExcelColumn;
import cn.emay.excel.common.ExcelSheet;
import cn.emay.excel.read.ExcelReadHelper;
import cn.emay.excel.read.handler.SheetReadHandler;

/**
 * 基于Schema的读取处理器
 * 
 * @author Frank
 *
 * @param <D>
 */
public class SheetReadBySchemaHandler<D extends Object> implements SheetReadHandler {

	/**
	 * 所有的列定义集合
	 */
	private Map<String, ExcelColumn> schemaMap;
	/**
	 * 所有的字段集合
	 */
	private Map<String, Field> fieldMap;
	/**
	 * 定义class
	 */
	private Class<D> schemaClass;
	/**
	 * 数据处理器
	 */
	private DataReader<D> dataHandler;
	/**
	 * 当前sheet页的编号
	 */
	private int sheetIndex;
	/**
	 * 当前sheet页的名字
	 */
	private String sheetName;
	/**
	 * 当前数据
	 */
	private D cur;
	/**
	 * 列对应的标题
	 */
	private Map<Integer, String> colTitles = new HashMap<>();
	/**
	 * 读取标题的行
	 */
	private int readTitleRowIndex;
	/**
	 * 读取数据的起始行
	 */
	private int readDataStartRowIndex;
	/**
	 * 读取数据的结束行
	 */
	private int readDataEndRowIndex;
	/**
	 * 是否按照列序号读取
	 */
	private boolean readByIndex;

	/**
	 * 
	 * @param dataHandler
	 *            数据处理器
	 */
	public SheetReadBySchemaHandler(DataReader<D> dataHandler) {
		if (dataHandler == null) {
			throw new IllegalArgumentException("dataHandler is null");
		}
		schemaClass = dataHandler.getShemaClass();
		if (schemaClass == null) {
			throw new IllegalArgumentException("schemaClass is null");
		}
		if (!schemaClass.isAnnotationPresent(ExcelSheet.class)) {
			throw new IllegalArgumentException("schemaClass[" + schemaClass.getName() + "] is not has Annotation : " + ExcelSheet.class.getName());
		}
		ExcelSheet schema = schemaClass.getAnnotation(ExcelSheet.class);
		Field[] fields = schemaClass.getDeclaredFields();
		if (fields == null || fields.length == 0) {
			throw new IllegalArgumentException("schemaClass[" + schemaClass.getName() + "] is not has  filed  ");
		}
		this.readByIndex = schema.readColumnBy().equalsIgnoreCase("Index");
		this.readTitleRowIndex = schema.readTitleRowIndex();
		this.readDataStartRowIndex = schema.readDataStartRowIndex();
		this.readDataEndRowIndex = schema.readDataEndRowIndex();
		if (readByIndex == false && readTitleRowIndex < 0) {
			throw new IllegalArgumentException("error : readColumnBy = Title and readTitleRowIndex < 0");
		}
		if (readByIndex == false && readDataStartRowIndex <= readTitleRowIndex) {
			throw new IllegalArgumentException("error : readDataStartRowIndex[" + readDataStartRowIndex + "] < readTitleRowIndex[" + readTitleRowIndex + "]");
		}
		if (readByIndex == false && readDataEndRowIndex <= readTitleRowIndex) {
			throw new IllegalArgumentException("error : readDataEndRowIndex[" + readDataEndRowIndex + "] < readTitleRowIndex[" + readTitleRowIndex + "]");
		}
		Map<String, ExcelColumn> schemaMap = new HashMap<>();
		Map<String, Field> fieldMap = new HashMap<>();
		Set<String> titles = new HashSet<>();
		Set<Integer> columnIndexs = new HashSet<>();
		for (Field field : fields) {
			if (!field.isAnnotationPresent(ExcelColumn.class)) {
				continue;
			}
			ExcelColumn csma = field.getAnnotation(ExcelColumn.class);
			if (columnIndexs.contains(csma.columnIndex())) {
				throw new IllegalArgumentException(schemaClass.getName() + " has same columnIndex filed.");
			}
			if (titles.contains(csma.title())) {
				throw new IllegalArgumentException(schemaClass.getName() + " has same title filed.");
			}
			titles.add(csma.title());
			columnIndexs.add(csma.columnIndex());
			if (readByIndex) {
				schemaMap.put(String.valueOf(csma.columnIndex()), csma);
				fieldMap.put(String.valueOf(csma.columnIndex()), field);
			} else {
				schemaMap.put(csma.title(), csma);
				fieldMap.put(csma.title(), field);
			}
			field.setAccessible(true);
		}
		if (schemaMap.size() == 0) {
			throw new IllegalArgumentException("schemaClass[" + schemaClass.getName() + "] is not has ExcelColumnSchema filed  ");
		}
		this.schemaMap = schemaMap;
		this.fieldMap = fieldMap;
		this.dataHandler = dataHandler;
	}

	@Override
	public int getStartReadRowIndex() {
		return 0;
	}

	@Override
	public int getEndReadRowIndex() {
		return readDataEndRowIndex;
	}

	@Override
	public void begin(int sheetIndex, String sheetName) {
		this.sheetIndex = sheetIndex;
		this.sheetName = sheetName;
	}

	@Override
	public void beginRow(int rowIndex) {
		if (rowIndex < readDataStartRowIndex) {
			return;
		}
		try {
			cur = schemaClass.newInstance();
		} catch (InstantiationException e) {
			throw new IllegalArgumentException(schemaClass.getName() + " can't be new Instance", e);
		} catch (IllegalAccessException e) {
			throw new IllegalArgumentException(schemaClass.getName() + " can't be new Instance", e);
		}
	}

	@Override
	public void handleXlsCell(int rowIndex, int columnIndex, Cell cell) {
		if (readByIndex == false && rowIndex == readTitleRowIndex) {
			String title = ExcelReadHelper.readString(cell);
			colTitles.put(columnIndex, title == null ? "" : title);
		}
		if (rowIndex < readDataStartRowIndex) {
			return;
		}
		if (cell == null) {
			return;
		}
		String key = null;
		if (readByIndex) {
			key = String.valueOf(columnIndex);
		} else {
			key = colTitles.get(columnIndex);
		}
		if (key == null) {
			return;
		}
		Field field = fieldMap.get(key);
		if (field == null) {
			return;
		}
		ExcelColumn columnSchema = schemaMap.get(key);
		if (columnSchema == null) {
			return;
		}
		Object obj = null;
		try {
			if (field.getType().isAssignableFrom(int.class) || field.getType().isAssignableFrom(Integer.class)) {
				obj = ExcelReadHelper.readInteger(cell);
			} else if (field.getType().isAssignableFrom(Double.class) || field.getType().isAssignableFrom(double.class)) {
				obj = ExcelReadHelper.readDouble(cell, parseNumberExpress(columnSchema.express()));
			} else if (field.getType().isAssignableFrom(Long.class) || field.getType().isAssignableFrom(long.class)) {
				obj = ExcelReadHelper.readLong(cell);
			} else if (field.getType().isAssignableFrom(BigDecimal.class)) {
				obj = ExcelReadHelper.readBigDecimal(cell, parseNumberExpress(columnSchema.express()));
			} else if (field.getType().isAssignableFrom(Date.class)) {
				obj = ExcelReadHelper.readDate(cell, columnSchema.express());
			} else if (field.getType().isAssignableFrom(Boolean.class) || field.getType().isAssignableFrom(boolean.class)) {
				obj = ExcelReadHelper.readBoolean(cell);
			} else if (field.getType().isAssignableFrom(String.class)) {
				obj = ExcelReadHelper.readString(cell);
			}
			if (obj != null) {
				field.set(cur, obj);
			}
		} catch (IllegalArgumentException e) {
			throw new IllegalArgumentException(
					"sheet(" + sheetName + "):[" + sheetIndex + "]-row[" + rowIndex + "]-column[" + columnIndex + "] read[" + obj + "] and set[" + field.getName() + "] error", e);
		} catch (IllegalAccessException e) {
			throw new IllegalArgumentException(
					"sheet(" + sheetName + "):[" + sheetIndex + "]-row[" + rowIndex + "]-column[" + columnIndex + "] read[" + obj + "] and set[" + field.getName() + "] error", e);
		}
	}

	@Override
	public void handleXlsxCell(int rowIndex, int columnIndex, int formatIndex, String value) {
		if (readByIndex == false && rowIndex == readTitleRowIndex) {
			String title = ExcelReadHelper.readString(value);
			colTitles.put(columnIndex, title == null ? "" : title);
		}
		if (rowIndex < readDataStartRowIndex) {
			return;
		}
		if (value == null) {
			return;
		}
		String key = null;
		if (readByIndex) {
			key = String.valueOf(columnIndex);
		} else {
			key = colTitles.get(columnIndex);
		}
		if (key == null) {
			return;
		}
		Field field = fieldMap.get(key);
		if (field == null) {
			return;
		}
		ExcelColumn columnSchema = schemaMap.get(key);
		if (columnSchema == null) {
			return;
		}
		Object obj = null;
		try {
			if (field.getType().isAssignableFrom(int.class) || field.getType().isAssignableFrom(Integer.class)) {
				obj = ExcelReadHelper.readInteger(value);
			} else if (field.getType().isAssignableFrom(Double.class) || field.getType().isAssignableFrom(double.class)) {
				obj = ExcelReadHelper.readDouble(value, parseNumberExpress(columnSchema.express()));
			} else if (field.getType().isAssignableFrom(Long.class) || field.getType().isAssignableFrom(long.class)) {
				obj = ExcelReadHelper.readLong(value);
			} else if (field.getType().isAssignableFrom(BigDecimal.class)) {
				obj = ExcelReadHelper.readBigDecimal(value, parseNumberExpress(columnSchema.express()));
			} else if (field.getType().isAssignableFrom(Date.class)) {
				obj = ExcelReadHelper.readDate(formatIndex, value, columnSchema.express());
			} else if (field.getType().isAssignableFrom(Boolean.class) || field.getType().isAssignableFrom(boolean.class)) {
				obj = ExcelReadHelper.readBoolean(value);
			} else if (field.getType().isAssignableFrom(String.class)) {
				obj = ExcelReadHelper.readString(value);
			}
			if (obj != null) {
				field.set(cur, obj);
			}
		} catch (IllegalArgumentException e) {
			throw new IllegalArgumentException(
					"sheet(" + sheetName + "):[" + sheetIndex + "]-row[" + rowIndex + "]-column[" + columnIndex + "] read[" + obj + "] and set[" + field.getName() + "] error", e);
		} catch (IllegalAccessException e) {
			throw new IllegalArgumentException(
					"sheet(" + sheetName + "):[" + sheetIndex + "]-row[" + rowIndex + "]-column[" + columnIndex + "] read[" + obj + "] and set[" + field.getName() + "] error", e);
		}
	}

	@Override
	public void endRow(int rowIndex) {
		if (cur != null) {
			dataHandler.handlerRowData(rowIndex, cur);
		}
	}

	@Override
	public void end(int sheetIndex, String sheetName) {

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
};
