package cn.emay.excel.read.reader.impl;

import cn.emay.excel.common.schema.base.ColumnSchema;
import cn.emay.excel.common.schema.base.SheetReadSchemaParams;
import cn.emay.excel.common.schema.base.SheetSchema;
import cn.emay.excel.read.handler.SchemaSheetDataHandler;
import cn.emay.excel.read.handler.SheetDataHandler;
import cn.emay.excel.read.reader.SheetReader;
import cn.emay.excel.utils.ExcelReadUtils;
import cn.emay.excel.utils.ExcelUtils;
import cn.emay.utils.clazz.ClassUtils;
import org.apache.poi.ss.usermodel.Cell;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

/**
 * 定义方式读取
 *
 * @param <D> 数据
 * @author Frank
 */
public class SchemaSheetReader<D> implements SheetReader {

    /**
     * 定义
     */
    private final SheetSchema schema;
    /**
     * 读参数集
     */
    private final SheetReadSchemaParams readSchemaParams;
    /**
     * 数据处理器
     */
    private final SheetDataHandler<D> dataReader;
    /**
     * 是否按照列序号读取
     */
    private final boolean readByIndex;

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
    private final Map<String, Field> fields = new HashMap<>();
    /**
     * 列对应的标题
     */
    private final Map<Integer, String> colTitles = new HashMap<>();

    /**
     * @param schemaSheetDataHandler 基于定义的数据处理器
     */
    public SchemaSheetReader(SchemaSheetDataHandler<D> schemaSheetDataHandler) {
        this(schemaSheetDataHandler.getSheetSchema(), schemaSheetDataHandler);
    }

    /**
     * @param schema     读取定义
     * @param dataReader 数据处理器
     */
    public SchemaSheetReader(SheetSchema schema, SheetDataHandler<D> dataReader) {
        if (schema == null) {
            throw new IllegalArgumentException("schema is null");
        }
        schema.checkRead();
        if (dataReader == null) {
            throw new IllegalArgumentException("dataReader is null");
        }
        this.schema = schema;
        this.dataReader = dataReader;
        this.readSchemaParams = schema.getSheetReadSchemaParams();
        this.readByIndex = readSchemaParams.readByIndex();
        Set<String> titles = new HashSet<>();
        Set<Integer> indexs = new HashSet<>();
        Field[] fieldArray = ClassUtils.getAllFields(dataReader.getDataClass());
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
        return this.readSchemaParams.getReadDataEndRowIndex();
    }

    @Override
    public void begin(int sheetIndex, String sheetName) {
        this.curSheetIndex = sheetIndex;
        this.curSheetName = sheetName;
    }

    @Override
    public void beginRow(int rowIndex) {
        if (rowIndex < this.readSchemaParams.getReadDataStartRowIndex()) {
            return;
        }
        curData = ExcelUtils.newData(dataReader.getDataClass());
    }

    @Override
    public void handleXlsCell(int rowIndex, int columnIndex, Cell cell) {
        if (!readByIndex && rowIndex == this.readSchemaParams.getReadTitleRowIndex()) {
            String title = ExcelReadUtils.readString(cell);
            colTitles.put(columnIndex, title == null ? "" : title);
        }
        if (rowIndex < this.readSchemaParams.getReadDataStartRowIndex()) {
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
            obj = ExcelReadUtils.read(field.getType(), cell, columnSchema.getExpress());
            if (obj != null) {
                field.set(curData, obj);
            }
        } catch (Exception e) {
            throw new IllegalArgumentException(
                    "sheet(" + curSheetName + "):[" + curSheetIndex + "]-row[" + rowIndex + "]-column[" + columnIndex + "] read[" + obj + "] and set[" + field.getName() + "] error", e);
        }
    }

    @Override
    public void handleXlsxCell(int rowIndex, int columnIndex, String value) {
        if (!readByIndex && rowIndex == this.readSchemaParams.getReadTitleRowIndex()) {
            String title = ExcelReadUtils.readString(value);
            colTitles.put(columnIndex, title == null ? "" : title);
        }
        if (rowIndex < this.readSchemaParams.getReadDataStartRowIndex()) {
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
            obj = ExcelReadUtils.read(field.getType(), value, columnSchema.getExpress());
            if (obj != null) {
                field.set(curData, obj);
            }
        } catch (Exception e) {
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