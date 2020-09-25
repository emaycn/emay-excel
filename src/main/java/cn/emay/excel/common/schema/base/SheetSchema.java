package cn.emay.excel.common.schema.base;

import cn.emay.excel.common.schema.annotation.ExcelColumn;
import cn.emay.excel.common.schema.annotation.ExcelSheet;
import cn.emay.utils.clazz.ClassUtils;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.Map;

/**
 * 表定义
 *
 * @author Frank
 */
public class SheetSchema {

    /**
     * 表定义的参数集
     */
    private SheetWriteSchemaParams writeSchemaParams;
    /**
     * 表定义的参数集
     */
    private SheetReadSchemaParams readSchemaParams;
    /**
     * 列定义，根据字段名匹配
     */
    private final Map<String, ColumnSchema> columnSchemas = new HashMap<>();

    /**
     * @param dataClass 数据Class
     */
    public SheetSchema(Class<?> dataClass) {
        if (dataClass == null) {
            throw new IllegalArgumentException("dataClass is null");
        }
        if (dataClass.isAnnotationPresent(ExcelSheet.class)) {
            ExcelSheet sheet = dataClass.getAnnotation(ExcelSheet.class);
            this.setSheetSchemaParams(sheet);
            Field[] fields = ClassUtils.getAllFields(dataClass);
            for (Field field : fields) {
                if (field.isAnnotationPresent(ExcelColumn.class)) {
                    this.setColumnSchema(field.getName(), field.getAnnotation(ExcelColumn.class));
                }
            }
        }
    }

    /**
     * @param writeSchemaParams        表定义写参数
     * @param columnSchemaByFieldNames 列定义，根据字段名匹配
     */
    public SheetSchema(SheetWriteSchemaParams writeSchemaParams, Map<String, ColumnSchema> columnSchemaByFieldNames) {
        this(writeSchemaParams, null, columnSchemaByFieldNames);
    }

    /**
     * @param readSchemaParams         表定义读参数
     * @param columnSchemaByFieldNames 列定义，根据字段名匹配
     */
    public SheetSchema(SheetReadSchemaParams readSchemaParams, Map<String, ColumnSchema> columnSchemaByFieldNames) {
        this(null, readSchemaParams, columnSchemaByFieldNames);
    }

    /**
     * @param writeSchemaParams        表定义写参数
     * @param readSchemaParams         表定义读参数
     * @param columnSchemaByFieldNames 列定义，根据字段名匹配
     */
    public SheetSchema(SheetWriteSchemaParams writeSchemaParams, SheetReadSchemaParams readSchemaParams, Map<String, ColumnSchema> columnSchemaByFieldNames) {
        this.writeSchemaParams = writeSchemaParams;
        this.readSchemaParams = readSchemaParams;
        this.setColumnSchemas(columnSchemaByFieldNames);
    }

    /**
     * 传入表定义写参数集
     *
     * @param writeSchemaParams 表定义写参数集
     */
    public void setSheetWriteSchemaParams(SheetWriteSchemaParams writeSchemaParams) {
        this.writeSchemaParams = writeSchemaParams;
    }

    /**
     * 传入表定义读参数集
     *
     * @param readSchemaParams 表定义读参数集
     */
    public void setSheetWriteSchemaParams(SheetReadSchemaParams readSchemaParams) {
        this.readSchemaParams = readSchemaParams;
    }

    /**
     * 传入表定义
     *
     * @param sheet 表定义注解
     */
    public void setSheetSchemaParams(ExcelSheet sheet) {
        this.writeSchemaParams = new SheetWriteSchemaParams();
        this.writeSchemaParams.setAutoWidth(sheet.isAutoWidth());
        this.writeSchemaParams.setAutoWrap(sheet.isAutoWrap());
        this.writeSchemaParams.setCacheNumber(sheet.cacheNumber());
        this.writeSchemaParams.setContentRgbColor(sheet.contentRgbColor());
        this.writeSchemaParams.setNeedBorder(sheet.isNeedBorder());
        this.writeSchemaParams.setTitleRgbColor(sheet.titleRgbColor());
        this.writeSchemaParams.setWriteSheetName(sheet.writeSheetName());
        this.writeSchemaParams.setWriteTile(sheet.isWriteTile());
        this.readSchemaParams = new SheetReadSchemaParams();
        this.readSchemaParams.setReadColumnBy(sheet.readColumnBy());
        this.readSchemaParams.setReadDataEndRowIndex(sheet.readDataEndRowIndex());
        this.readSchemaParams.setReadDataStartRowIndex(sheet.readDataStartRowIndex());
        this.readSchemaParams.setReadTitleRowIndex(sheet.readTitleRowIndex());
    }

    /**
     * 传入对应字段的列定义集合
     *
     * @param columnSchemaByFieldNames 以字段名为key的列定义集合
     */
    public void setColumnSchemas(Map<String, ColumnSchema> columnSchemaByFieldNames) {
        if (columnSchemaByFieldNames == null) {
            return;
        }
        this.columnSchemas.putAll(columnSchemaByFieldNames);
    }

    /**
     * 传入对应字段的列定义对象
     *
     * @param fieldName    字段名
     * @param columnSchema 列定义
     */
    public void setColumnSchema(String fieldName, ColumnSchema columnSchema) {
        if (fieldName == null) {
            return;
        }
        this.columnSchemas.put(fieldName, columnSchema);
    }

    /**
     * 传入对应字段的列定义注解
     *
     * @param fieldName   字段名
     * @param excelColumn 列定义注解
     */
    public void setColumnSchema(String fieldName, ExcelColumn excelColumn) {
        this.setColumnSchema(fieldName, new ColumnSchema(excelColumn.index(), excelColumn.title(), excelColumn.express()));
    }

    /**
     * 获取表定义写参数集
     *
     * @return 表定义写参数集
     */
    public SheetWriteSchemaParams getSheetWriteSchemaParams() {
        return this.writeSchemaParams;
    }

    /**
     * 获取表定义读参数集
     *
     * @return 表定义读参数集
     */
    public SheetReadSchemaParams getSheetReadSchemaParams() {
        return this.readSchemaParams;
    }

    /**
     * 根据字段名获取列定义
     *
     * @param fieldName 字段名
     * @return 列定义
     */
    public ColumnSchema getExcelColumnByFieldName(String fieldName) {
        if (fieldName == null) {
            return null;
        }
        return this.columnSchemas.get(fieldName);
    }

    /**
     * 检测定义正确性
     */
    public void checkWrite() {
        if (this.columnSchemas.size() == 0) {
            throw new IllegalArgumentException("has not Column for field");
        }
        if (this.writeSchemaParams == null) {
            throw new IllegalArgumentException("sheetSchema is null");
        }
    }

    /**
     * 检测定义正确性
     */
    public void checkRead() {
        if (this.columnSchemas.size() == 0) {
            throw new IllegalArgumentException("has not Column for field");
        }
        if (this.readSchemaParams == null) {
            throw new IllegalArgumentException("readSchemaParams is null");
        }
        boolean readByIndex = this.readSchemaParams.readByIndex();
        int readTitleRowIndex = this.readSchemaParams.getReadTitleRowIndex();
        int readDataStartRowIndex = this.readSchemaParams.getReadDataStartRowIndex();
        int readDataEndRowIndex = this.readSchemaParams.getReadDataEndRowIndex();
        if (!readByIndex && readTitleRowIndex < 0) {
            throw new IllegalArgumentException("sheetSchemaParams's readColumnBy = Title and readTitleRowIndex < 0");
        }
        if (!readByIndex && readDataStartRowIndex <= readTitleRowIndex) {
            throw new IllegalArgumentException("sheetSchemaParams's readDataStartRowIndex[" + readDataStartRowIndex + "] < readTitleRowIndex[" + readTitleRowIndex + "]");
        }
        if (!readByIndex && readDataEndRowIndex <= readTitleRowIndex) {
            throw new IllegalArgumentException("sheetSchemaParams's readDataEndRowIndex[" + readDataEndRowIndex + "] < readTitleRowIndex[" + readTitleRowIndex + "]");
        }
    }
}
