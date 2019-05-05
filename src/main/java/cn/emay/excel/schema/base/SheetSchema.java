package cn.emay.excel.schema.base;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.Map;

import cn.emay.excel.schema.annotation.ExcelColumn;
import cn.emay.excel.schema.annotation.ExcelSheet;
import cn.emay.excel.schema.base.ColumnSchema;
import cn.emay.excel.schema.base.SheetSchemaParams;

/**
 * Excel表定义
 * 
 * @author Frank
 *
 * @param <D>
 *            数据类型
 */
public class SheetSchema<D> {

	/**
	 * 数据Class
	 */
	private Class<D> dataClass;
	/**
	 * 表定义
	 */
	private SheetSchemaParams sheetSchemaParams;
	/**
	 * 列定义，根据字段名匹配
	 */
	private Map<String, ColumnSchema> columnSchemas = new HashMap<>();

	/**
	 * 
	 * @param dataClass
	 *            数据Class
	 */
	public SheetSchema(Class<D> dataClass) {
		if (dataClass == null) {
			throw new IllegalArgumentException("dataClass is null");
		}
		this.dataClass = dataClass;
		if (dataClass.isAnnotationPresent(ExcelSheet.class)) {
			ExcelSheet sheet = dataClass.getAnnotation(ExcelSheet.class);
			this.setSheetSchemaParams(sheet);
			Field[] fields = dataClass.getDeclaredFields();
			for (Field field : fields) {
				if (field.isAnnotationPresent(ExcelColumn.class)) {
					this.setColumnSchema(field.getName(), field.getAnnotation(ExcelColumn.class));
				}
			}
		}
		this.dataClass = dataClass;
	}

	/**
	 * 
	 * @param dataClass
	 *            数据Class
	 * @param sheetSchemaParams
	 *            表定义参数
	 * @param columnSchemaByFieldNames
	 *            列定义，根据字段名匹配
	 */
	public SheetSchema(Class<D> dataClass, SheetSchemaParams sheetSchemaParams, Map<String, ColumnSchema> columnSchemaByFieldNames) {
		this(dataClass);
		this.setSheetSchemaParams(sheetSchemaParams);
		this.columnSchemas.putAll(columnSchemaByFieldNames);
	}

	/**
	 * 传入表定义
	 * 
	 * @param sheetSchemaParams
	 *            表定义对象
	 */
	public void setSheetSchemaParams(SheetSchemaParams sheetSchemaParams) {
		this.sheetSchemaParams = sheetSchemaParams;
	}

	/**
	 * 传入表定义
	 * 
	 * @param sheet
	 *            表定义注解
	 */
	public void setSheetSchemaParams(ExcelSheet sheet) {
		this.sheetSchemaParams = new SheetSchemaParams();
		this.sheetSchemaParams.setAutoWidth(sheet.isAutoWidth());
		this.sheetSchemaParams.setAutoWrap(sheet.isAutoWrap());
		this.sheetSchemaParams.setCacheNumber(sheet.cacheNumber());
		this.sheetSchemaParams.setContentRgbColor(sheet.contentRgbColor());
		this.sheetSchemaParams.setNeedBorder(sheet.isNeedBorder());
		this.sheetSchemaParams.setReadColumnBy(sheet.readColumnBy());
		this.sheetSchemaParams.setReadDataEndRowIndex(sheet.readDataEndRowIndex());
		this.sheetSchemaParams.setReadDataStartRowIndex(sheet.readDataStartRowIndex());
		this.sheetSchemaParams.setReadTitleRowIndex(sheet.readTitleRowIndex());
		this.sheetSchemaParams.setTitleRgbColor(sheet.titleRgbColor());
		this.sheetSchemaParams.setWriteSheetName(sheet.writeSheetName());
		this.sheetSchemaParams.setWriteTile(sheet.isWriteTile());
	}

	/**
	 * 传入对应字段的列定义对象
	 * 
	 * @param fieldName
	 *            字段名
	 * @param columnSchema
	 *            列定义
	 */
	public void setColumnSchema(String fieldName, ColumnSchema columnSchema) {
		this.columnSchemas.put(fieldName, columnSchema);
	}

	/**
	 * 传入对应字段的列定义注解
	 * 
	 * @param fieldName
	 *            字段名
	 * @param excelColumn
	 *            列定义注解
	 */
	public void setColumnSchema(String fieldName, ExcelColumn excelColumn) {
		this.setColumnSchema(fieldName, new ColumnSchema(excelColumn.index(), excelColumn.title(), excelColumn.express()));
	}

	/**
	 * 获取表定义
	 * 
	 * @return
	 */
	public SheetSchemaParams getSheetSchemaParams() {
		return this.sheetSchemaParams;
	}

	/**
	 * 根据字段名获取列定义
	 * 
	 * @param fieldName
	 *            字段名
	 * @return
	 */
	public ColumnSchema getExcelColumnByFieldName(String fieldName) {
		return this.columnSchemas.get(fieldName);
	}

	/**
	 * 获取数据class
	 * 
	 * @return
	 */
	public Class<D> getDataClass() {
		return dataClass;
	}

	/**
	 * 检测定义正确性
	 */
	public void check() {
		if (this.sheetSchemaParams == null) {
			throw new IllegalArgumentException("sheetSchema is null");
		}
		if (this.columnSchemas.size() == 0) {
			throw new IllegalArgumentException("has not Column for field");
		}
		boolean readByIndex = this.sheetSchemaParams.readByIndex();
		int readTitleRowIndex = this.sheetSchemaParams.getReadTitleRowIndex();
		int readDataStartRowIndex = this.sheetSchemaParams.getReadDataStartRowIndex();
		int readDataEndRowIndex = this.sheetSchemaParams.getReadDataEndRowIndex();
		if (readByIndex == false && readTitleRowIndex < 0) {
			throw new IllegalArgumentException("sheetSchemaParams's readColumnBy = Title and readTitleRowIndex < 0");
		}
		if (readByIndex == false && readDataStartRowIndex <= readTitleRowIndex) {
			throw new IllegalArgumentException("sheetSchemaParams's readDataStartRowIndex[" + readDataStartRowIndex + "] < readTitleRowIndex[" + readTitleRowIndex + "]");
		}
		if (readByIndex == false && readDataEndRowIndex <= readTitleRowIndex) {
			throw new IllegalArgumentException("sheetSchemaParams's readDataEndRowIndex[" + readDataEndRowIndex + "] < readTitleRowIndex[" + readTitleRowIndex + "]");
		}
	}

	/**
	 * 新建一个数据实例
	 * 
	 * @return
	 */
	public D newData() {
		try {
			return getDataClass().newInstance();
		} catch (InstantiationException e) {
			throw new IllegalArgumentException(getDataClass().getName() + " can't be new Instance", e);
		} catch (IllegalAccessException e) {
			throw new IllegalArgumentException(getDataClass().getName() + " can't be new Instance", e);
		}
	}

}
