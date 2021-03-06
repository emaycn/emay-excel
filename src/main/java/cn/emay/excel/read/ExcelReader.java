package cn.emay.excel.read;

import cn.emay.excel.common.ExcelVersion;
import cn.emay.excel.common.schema.base.SheetSchema;
import cn.emay.excel.read.core.BaseReader;
import cn.emay.excel.read.core.XlsReader;
import cn.emay.excel.read.core.XlsxReader;
import cn.emay.excel.read.handler.SchemaSheetDataHandler;
import cn.emay.excel.read.handler.SheetDataHandler;
import cn.emay.excel.read.reader.SheetReader;
import cn.emay.excel.read.reader.impl.SchemaSheetReader;
import cn.emay.excel.utils.ExcelReadUtils;
import cn.emay.excel.utils.ExcelUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

/**
 * Excel基础读取<br/>
 * XLSX统一采用SAX方式读取
 *
 * @author Frank
 */
public class ExcelReader {

    /**
     * 从Excel文件中读取第一个表格<br/>
     * dataClass实现了@ExcelSheet注解,其字段实现了@ExcelColumn注解
     *
     * @param excelPath 路径
     * @param dataClass 数据Class
     * @return 数据
     */
    public static <D> List<D> readFirstSheet(String excelPath, Class<D> dataClass) {
        return readBySheetIndex(excelPath, 0, dataClass);
    }

    /**
     * 从Excel文件中读取一个表格<br/>
     * dataClass实现了@ExcelSheet注解,其字段实现了@ExcelColumn注解
     *
     * @param excelPath 路径
     * @param sheetName Sheet页名字
     * @param dataClass 数据Class
     * @return 数据
     */
    public static <D> List<D> readBySheetName(String excelPath, String sheetName, Class<D> dataClass) {
        ReturnSchemaDataReader<D> dataReader = new ReturnSchemaDataReader<>(dataClass);
        readBySheetName(excelPath, sheetName, dataReader);
        return dataReader.getResult();
    }

    /**
     * 从Excel文件中读取一个表格<br/>
     * dataClass实现了@ExcelSheet注解,其字段实现了@ExcelColumn注解
     *
     * @param excelPath  路径
     * @param sheetIndex Sheet Index
     * @param dataClass  数据Class
     * @return 数据
     */
    public static <D> List<D> readBySheetIndex(String excelPath, int sheetIndex, Class<D> dataClass) {
        ReturnSchemaDataReader<D> dataReader = new ReturnSchemaDataReader<>(dataClass);
        readBySheetIndex(excelPath, sheetIndex, dataReader);
        return dataReader.getResult();
    }

    /**
     * 从Excel文件中读取第一个表格<br/>
     * dataClass 实现了@ExcelSheet注解,其字段实现了@ExcelColumn注解
     *
     * @param excelPath   路径
     * @param dataHandler 数据处理器
     */
    public static <D> void readFirstSheet(String excelPath, SheetDataHandler<D> dataHandler) {
        readBySheetIndex(excelPath, 0, dataHandler);
    }

    /**
     * 从Excel文件中读取一个表格<br/>
     * dataClass实现了@ExcelSheet注解,其字段实现了@ExcelColumn注解
     *
     * @param excelPath   路径
     * @param sheetName   Sheet页名字
     * @param dataHandler 数据处理器
     */
    public static <D> void readBySheetName(String excelPath, String sheetName, SheetDataHandler<D> dataHandler) {
        SchemaSheetReader<D> handler = new SchemaSheetReader<>(new SheetSchema(dataHandler.getDataClass()), dataHandler);
        readBySheetName(excelPath, sheetName, handler);
    }

    /**
     * 从文件中读取Excel表格<br/>
     * 按照表名匹配读取处理器<br/>
     *
     * @param excelPath      路径
     * @param handlersByName 按照表名匹配的Sheet读取表格定义读取器
     */
    @SuppressWarnings({"rawtypes", "unchecked"})
    public static void readBySheetNamesWithHandler(String excelPath, Map<String, SheetDataHandler<?>> handlersByName) {
        if (handlersByName == null) {
            return;
        }
        Map<String, SheetReader> readers = new HashMap<>(handlersByName.size());
        for (String name : handlersByName.keySet()) {
            SheetDataHandler<?> handler = handlersByName.get(name);
            SchemaSheetReader<?> reader = new SchemaSheetReader(new SheetSchema(handler.getDataClass()), handler);
            readers.put(name, reader);
        }
        readBySheetNames(excelPath, readers);
    }

    /**
     * 从Excel文件中读取一个表格<br/>
     * dataClass实现了@ExcelSheet注解,其字段实现了@ExcelColumn注解
     *
     * @param excelPath   路径
     * @param sheetIndex  Sheet Index
     * @param dataHandler 数据处理器
     */
    public static <D> void readBySheetIndex(String excelPath, int sheetIndex, SheetDataHandler<D> dataHandler) {
        SchemaSheetReader<D> handler = new SchemaSheetReader<>(new SheetSchema(dataHandler.getDataClass()), dataHandler);
        readBySheetIndex(excelPath, sheetIndex, handler);
    }

    /**
     * 从文件中读取Excel表格<br/>
     * 按照表序号匹配读取处理器<br/>
     *
     * @param excelPath       路径
     * @param handlersByIndex 按照Index匹配的Sheet读取表格定义读取器
     */
    @SuppressWarnings({"rawtypes", "unchecked"})
    public static void readBySheetIndexsWithHandler(String excelPath, Map<Integer, SheetDataHandler<?>> handlersByIndex) {
        if (handlersByIndex == null) {
            return;
        }
        Map<Integer, SheetReader> readers = new HashMap<>(handlersByIndex.size());
        for (Integer index : handlersByIndex.keySet()) {
            SheetDataHandler<?> handler = handlersByIndex.get(index);
            SchemaSheetReader<?> reader = new SchemaSheetReader(new SheetSchema(handler.getDataClass()), handler);
            readers.put(index, reader);
        }
        readBySheetIndexs(excelPath, readers);
    }

    /**
     * 从文件中读取Excel表格<br/>
     *
     * @param excelPath 路径
     * @param handlers  Excel表格定义读取器(handlers顺序号即为读取ExccelSheet的编号)
     */
    @SuppressWarnings({"rawtypes", "unchecked"})
    public static void readByOrderWithSchema(String excelPath, SheetDataHandler<?>... handlers) {
        if (handlers == null) {
            return;
        }
        SheetReader[] readers = new SheetReader[handlers.length];
        for (int i = 0; i < handlers.length; i++) {
            SheetDataHandler<?> handler = handlers[i];
            readers[i] = new SchemaSheetReader(new SheetSchema(handler.getDataClass()), handler);
        }
        readByOrder(excelPath, readers);
    }

    /**
     * 从Excel文件中读取第一个表格<br/>
     *
     * @param excelPath              路径
     * @param schemaSheetDataHandler 表格定义读取器
     */
    public static <D> void readFirstSheet(String excelPath, SchemaSheetDataHandler<D> schemaSheetDataHandler) {
        readBySheetIndex(excelPath, 0, schemaSheetDataHandler);
    }

    /**
     * 从文件中读取Excel表格<br/>
     * 按照表序号匹配读取处理器<br/>
     *
     * @param excelPath       路径
     * @param handlersByIndex 按照Index匹配的Sheet读取表格定义读取器
     */
    @SuppressWarnings({"rawtypes", "unchecked"})
    public static void readBySheetIndexsWithSchema(String excelPath, Map<Integer, SchemaSheetDataHandler<?>> handlersByIndex) {
        if (handlersByIndex == null) {
            return;
        }
        Map<Integer, SheetReader> readers = new HashMap<>(handlersByIndex.size());
        for (Integer index : handlersByIndex.keySet()) {
            SchemaSheetDataHandler<?> schemaSheetDataHandler = handlersByIndex.get(index);
            SchemaSheetReader<?> handler = new SchemaSheetReader(schemaSheetDataHandler.getSheetSchema(), schemaSheetDataHandler);
            readers.put(index, handler);
        }
        readBySheetIndexs(excelPath, readers);
    }

    /**
     * 从Excel文件中读取一个表格<br/>
     *
     * @param excelPath              路径
     * @param sheetIndex             Sheet Index
     * @param schemaSheetDataHandler 表格定义读取器
     */
    public static <D> void readBySheetIndex(String excelPath, int sheetIndex, SchemaSheetDataHandler<D> schemaSheetDataHandler) {
        SchemaSheetReader<D> handler = new SchemaSheetReader<>(schemaSheetDataHandler.getSheetSchema(), schemaSheetDataHandler);
        readBySheetIndex(excelPath, sheetIndex, handler);
    }

    /**
     * 从文件中读取Excel表格<br/>
     *
     * @param excelPath 路径
     * @param handlers  Excel表格定义读取器(handlers顺序号即为读取ExccelSheet的编号)
     */
    @SuppressWarnings({"rawtypes", "unchecked"})
    public static void readByOrderWithSchema(String excelPath, SchemaSheetDataHandler<?>... handlers) {
        if (handlers == null) {
            return;
        }
        SheetReader[] readers = new SheetReader[handlers.length];
        for (int i = 0; i < handlers.length; i++) {
            readers[i] = new SchemaSheetReader(handlers[i].getSheetSchema(), handlers[i]);
        }
        readByOrder(excelPath, readers);
    }

    /**
     * 从Excel文件中读取一个表格<br/>
     *
     * @param excelPath              路径
     * @param sheetName              Sheet页名字
     * @param schemaSheetDataHandler 表格定义读取器
     */
    public static <D> void readBySheetName(String excelPath, String sheetName, SchemaSheetDataHandler<D> schemaSheetDataHandler) {
        SchemaSheetReader<D> handler = new SchemaSheetReader<>(schemaSheetDataHandler.getSheetSchema(), schemaSheetDataHandler);
        readBySheetName(excelPath, sheetName, handler);
    }

    /**
     * 从文件中读取Excel表格<br/>
     * 按照表名匹配读取处理器<br/>
     *
     * @param excelPath      路径
     * @param handlersByName 按照表名匹配的Sheet读取表格定义读取器
     */
    @SuppressWarnings({"rawtypes", "unchecked"})
    public static void readBySheetNamesWithSchema(String excelPath, Map<String, SchemaSheetDataHandler<?>> handlersByName) {
        if (handlersByName == null) {
            return;
        }
        Map<String, SheetReader> readers = new HashMap<>(handlersByName.size());
        for (String name : handlersByName.keySet()) {
            SchemaSheetDataHandler<?> schemaSheetDataHandler = handlersByName.get(name);
            SchemaSheetReader<?> handler = new SchemaSheetReader(schemaSheetDataHandler.getSheetSchema(), schemaSheetDataHandler);
            readers.put(name, handler);
        }
        readBySheetNames(excelPath, readers);
    }

    /**
     * 从文件中读取Excel表格第一个sheet页<br/>
     *
     * @param excelPath 路径
     * @param reader    Sheet读取处理器[处理器实例不可复用]
     */
    public static void readFirstSheet(String excelPath, SheetReader reader) {
        readBySheetIndex(excelPath, 0, reader);
    }

    /**
     * 从文件中读取Excel表格<br/>
     * 按照表格序号读取<br/>
     *
     * @param excelPath  路径
     * @param sheetIndex Sheet Index
     * @param reader     Sheet读取处理器[处理器实例不可复用]
     */
    public static void readBySheetIndex(String excelPath, int sheetIndex, SheetReader reader) {
        ExcelVersion version = ExcelUtils.parserPath(excelPath);
        readBySheetIndex(new File(excelPath), version, sheetIndex, reader);
    }

    /**
     * 从文件中读取Excel表格<br/>
     * 按照表序号匹配读取处理器<br/>
     *
     * @param excelPath      路径
     * @param readersByIndex 按照Index匹配的Sheet读取处理器集合[处理器实例不可复用]
     */
    public static void readBySheetIndexs(String excelPath, Map<Integer, SheetReader> readersByIndex) {
        ExcelVersion version = ExcelUtils.parserPath(excelPath);
        readBySheetIndexs(new File(excelPath), version, readersByIndex);
    }

    /**
     * 从文件中读取Excel表格<br/>
     *
     * @param excelPath 路径
     * @param readers   Excel表处理器(handlers顺序号即为读取ExccelSheet的编号)
     */
    public static void readByOrder(String excelPath, SheetReader... readers) {
        ExcelVersion version = ExcelUtils.parserPath(excelPath);
        readByOrder(new File(excelPath), version, readers);
    }

    /**
     * 从文件中读取Excel表格<br/>
     * 按照表格名字读取<br/>
     *
     * @param excelPath 路径
     * @param sheetName Sheet页名字
     * @param reader    Sheet读取处理器[处理器实例不可复用]
     */
    public static void readBySheetName(String excelPath, String sheetName, SheetReader reader) {
        ExcelVersion version = ExcelUtils.parserPath(excelPath);
        readBySheetName(new File(excelPath), version, sheetName, reader);
    }

    /**
     * 从文件中读取Excel表格<br/>
     * 按照表名匹配读取处理器<br/>
     *
     * @param excelPath     路径
     * @param readersByName 按照表名匹配的Sheet读取处理器集合[处理器实例不可复用]
     */
    public static void readBySheetNames(String excelPath, Map<String, SheetReader> readersByName) {
        ExcelVersion version = ExcelUtils.parserPath(excelPath);
        readBySheetNames(new File(excelPath), version, readersByName);

    }

    /**
     * 从文件中读取Excel表格第一个sheet页<br/>
     *
     * @param is        输入流
     * @param version   版本
     * @param dataClass 数据类型
     */
    public static <D> List<D> readFirstSheet(InputStream is, ExcelVersion version, Class<D> dataClass) {
        ReturnSchemaDataReader<D> dataHandler = new ReturnSchemaDataReader<>(dataClass);
        SchemaSheetReader<D> handler = new SchemaSheetReader<>(new SheetSchema(dataHandler.getDataClass()), dataHandler);
        readFirstSheet(is, version, handler);
        return dataHandler.getResult();
    }

    /**
     * 从文件中读取Excel表格第一个sheet页<br/>
     *
     * @param is      输入流
     * @param version 版本
     * @param reader  Sheet读取处理器[处理器实例不可复用]
     */
    public static void readFirstSheet(InputStream is, ExcelVersion version, SheetReader reader) {
        readBySheetIndex(is, version, 0, reader);
    }

    /**
     * 从输入流中读取Excel表格<br/>
     * 按照表格序号读取<br/>
     *
     * @param is         输入流
     * @param version    版本
     * @param sheetIndex Sheet Index
     * @param reader     Sheet读取处理器[处理器实例不可复用]
     */
    public static void readBySheetIndex(InputStream is, ExcelVersion version, int sheetIndex, SheetReader reader) {
        getReader(version).readBySheetIndex(is, sheetIndex, reader);
    }

    /**
     * 从文件中读取Excel表格<br/>
     * 按照表格序号读取<br/>
     *
     * @param file       文件
     * @param version    版本
     * @param sheetIndex Sheet Index
     * @param reader     Sheet读取处理器[处理器实例不可复用]
     */
    public static void readBySheetIndex(File file, ExcelVersion version, int sheetIndex, SheetReader reader) {
        getReader(version).readBySheetIndex(file, sheetIndex, reader);
    }

    /**
     * 从输入流中读取Excel表格<br/>
     *
     * @param is      输入流
     * @param readers Excel表处理器(handlers顺序号即为读取ExccelSheet的编号)
     */
    public static void readByOrder(InputStream is, ExcelVersion version, SheetReader... readers) {
        getReader(version).readByOrder(is, readers);
    }

    /**
     * 从文件中读取Excel表格<br/>
     *
     * @param file    文件
     * @param readers Excel表处理器(handlers顺序号即为读取ExccelSheet的编号)
     */
    public static void readByOrder(File file, ExcelVersion version, SheetReader... readers) {
        getReader(version).readByOrder(file, readers);
    }

    /**
     * 从文件中读取Excel表格<br/>
     * 按照表序号匹配读取处理器<br/>
     *
     * @param file           文件
     * @param version        版本
     * @param readersByIndex 按照Index匹配的Sheet读取处理器集合[处理器实例不可复用]
     */
    public static void readBySheetIndexs(File file, ExcelVersion version, Map<Integer, SheetReader> readersByIndex) {
        getReader(version).readBySheetIndexs(file, readersByIndex);
    }

    /**
     * 从输入流中读取Excel表格<br/>
     * 按照表序号匹配读取处理器<br/>
     *
     * @param is             输入流
     * @param version        版本
     * @param readersByIndex 按照Index匹配的Sheet读取处理器集合[处理器实例不可复用]
     */
    public static void readBySheetIndexs(InputStream is, ExcelVersion version, Map<Integer, SheetReader> readersByIndex) {
        getReader(version).readBySheetIndexs(is, readersByIndex);
    }

    /**
     * 从输入流中读取Excel表格<br/>
     * 按照表格名字读取<br/>
     *
     * @param is        输入流
     * @param version   版本
     * @param sheetName Sheet页名字
     * @param reader    Sheet读取处理器[处理器实例不可复用]
     */
    public static void readBySheetName(InputStream is, ExcelVersion version, String sheetName, SheetReader reader) {
        getReader(version).readBySheetName(is, sheetName, reader);
    }

    /**
     * 从文件中读取Excel表格<br/>
     * 按照表格名字读取<br/>
     *
     * @param file      文件
     * @param version   版本
     * @param sheetName Sheet页名字
     * @param reader    Sheet读取处理器[处理器实例不可复用]
     */
    public static void readBySheetName(File file, ExcelVersion version, String sheetName, SheetReader reader) {
        getReader(version).readBySheetName(file, sheetName, reader);
    }

    /**
     * 从文件中读取Excel表格<br/>
     * 按照表名匹配读取处理器<br/>
     *
     * @param file          文件
     * @param version       版本
     * @param readersByName 按照表名匹配的Sheet读取处理器集合[处理器实例不可复用]
     */
    public static void readBySheetNames(File file, ExcelVersion version, Map<String, SheetReader> readersByName) {
        getReader(version).readBySheetNames(file, readersByName);
    }

    /**
     * 从输入流中读取Excel表格<br/>
     * 按照表名匹配读取处理器<br/>
     *
     * @param is            输入流
     * @param version       版本
     * @param readersByName 按照表名匹配的Sheet读取处理器集合[处理器实例不可复用]
     */
    public static void readBySheetNames(InputStream is, ExcelVersion version, Map<String, SheetReader> readersByName) {
        getReader(version).readBySheetNames(is, readersByName);
    }

    /**
     * 按照坐标读取数据(String类型)<br/>
     * 大数据量的Excel禁用
     *
     * @param excelPath   Excel路径
     * @param coordinates 坐标集合
     * @return 数据
     */
    public static List<String> readByCoordinate(String excelPath, int[]... coordinates) {
        return readByCoordinate(String.class, excelPath, coordinates);
    }

    /**
     * 按照坐标读取数据(String类型)<br/>
     * 大数据量的Excel禁用
     *
     * @param is          输入流
     * @param version     Excel版本
     * @param coordinates 坐标集合
     * @return 数据
     */
    public static List<String> readByCoordinate(InputStream is, ExcelVersion version, int[]... coordinates) {
        return readByCoordinate(String.class, is, version, coordinates);
    }

    /**
     * 按照坐标读取数据<br/>
     * 大数据量的Excel禁用
     *
     * @param dataClass   读取的数据类型（所有数据都是此类型）
     * @param is          输入流
     * @param version     Excel版本
     * @param coordinates 坐标集合
     * @return 数据
     */
    public static <T> List<T> readByCoordinate(Class<T> dataClass, InputStream is, ExcelVersion version, int[]... coordinates) {
        if (ExcelVersion.XLS.equals(version)) {
            try (Workbook workbook = new HSSFWorkbook(is)) {
                return readByCoordinate(dataClass, workbook, coordinates);
            } catch (IOException e) {
                throw new IllegalArgumentException(e);
            }
        } else {
            try (Workbook workbook = new XSSFWorkbook(is)) {
                return readByCoordinate(dataClass, workbook, coordinates);
            } catch (IOException e) {
                throw new IllegalArgumentException(e);
            }
        }
    }

    /**
     * 按照坐标读取数据<br/>
     * 大数据量的Excel禁用
     *
     * @param dataClass   读取的数据类型（所有数据都是此类型）
     * @param excelPath   excel文件路径
     * @param coordinates 坐标集合
     * @return 数据
     */
    public static <T> List<T> readByCoordinate(Class<T> dataClass, String excelPath, int[]... coordinates) {
        ExcelVersion version = ExcelUtils.parserPath(excelPath);
        if (ExcelVersion.XLS.equals(version)) {
            try (
                    InputStream is = new FileInputStream(new File(excelPath));
                    Workbook workbook = new HSSFWorkbook(is)
            ) {
                return readByCoordinate(dataClass, workbook, coordinates);
            } catch (IOException e) {
                throw new IllegalArgumentException(e);
            }
        } else {
            try (Workbook workbook = new XSSFWorkbook(excelPath)) {
                return readByCoordinate(dataClass, workbook, coordinates);
            } catch (IOException e) {
                throw new IllegalArgumentException(e);
            }
        }
    }

    /**
     * 按照坐标读取数据<br/>
     * 大数据量的Excel禁用
     *
     * @param dataClass   读取的数据类型（所有数据都是此类型）
     * @param workbook    excel工作薄
     * @param coordinates 坐标集合
     * @return 数据
     */
    public static <T> List<T> readByCoordinate(Class<T> dataClass, Workbook workbook, int[]... coordinates) {
        if (coordinates == null) {
            throw new IllegalArgumentException("coordinates is null");
        }
        List<T> datas = new ArrayList<>(coordinates.length);
        Map<Integer, Map<Integer, Map<Integer, String>>> dataByCoordinate = new TreeMap<>();
        for (int[] coordinate : coordinates) {
            if (coordinate == null) {
                continue;
            }
            if (coordinate.length < 3) {
                continue;
            }
            int sheetIndex = coordinate[0];
            int rowIndex = coordinate[1];
            int columnIndex = coordinate[2];
            if (sheetIndex < 0) {
                continue;
            }
            if (rowIndex < 0) {
                continue;
            }
            if (columnIndex < 0) {
                continue;
            }
            Map<Integer, Map<Integer, String>> oneSheetMap = dataByCoordinate.computeIfAbsent(sheetIndex, k -> new TreeMap<>());
            Map<Integer, String> oneRowMap = oneSheetMap.computeIfAbsent(rowIndex, k -> new TreeMap<>());
            oneRowMap.put(columnIndex, null);
        }
        try (Workbook workbookNew = workbook) {
            int maxSheetIndex = workbookNew.getNumberOfSheets() - 1;
            for (Integer sheetIndex : dataByCoordinate.keySet()) {
                if (sheetIndex > maxSheetIndex) {
                    continue;
                }
                Sheet sheet = workbookNew.getSheetAt(sheetIndex);
                if (sheet == null) {
                    continue;
                }
                Map<Integer, Map<Integer, String>> oneSheetMap = dataByCoordinate.get(sheetIndex);
                for (Integer rowIndex : oneSheetMap.keySet()) {
                    Row row = sheet.getRow(rowIndex);
                    if (row == null) {
                        continue;
                    }
                    Map<Integer, String> oneRowMap = oneSheetMap.get(rowIndex);
                    for (Integer columnIndex : oneRowMap.keySet()) {
                        Cell cell = row.getCell(columnIndex);
                        if (cell == null) {
                            continue;
                        }
                        T data = ExcelReadUtils.read(dataClass, cell, null);
                        datas.add(data);
                    }
                }
            }
        } catch (IOException e) {
            throw new IllegalArgumentException(e);
        } finally {
            if (workbook != null) {
                if (SXSSFWorkbook.class.isAssignableFrom(workbook.getClass())) {
                    ((SXSSFWorkbook) workbook).dispose();
                }
            }
        }
        return datas;
    }

    /**
     * 根据版本获取Excel读取器
     *
     * @param version Excel版本
     * @return Excel读取器
     */
    private static BaseReader getReader(ExcelVersion version) {
        if (version == null) {
            return new XlsxReader();
        } else if (ExcelVersion.XLS.equals(version)) {
            return new XlsReader();
        } else {
            return new XlsxReader();
        }
    }

}

/**
 * 结果全部保存到内存中统一返回的数据处理器
 *
 * @param <D>
 * @author Frank
 */
class ReturnSchemaDataReader<D> implements SheetDataHandler<D> {

    /**
     * 所有数据
     */
    private final List<D> list = new ArrayList<>();
    /**
     * 数据Class
     */
    private final Class<D> dataClass;

    /**
     * @param dataClass 数据Class
     */
    public ReturnSchemaDataReader(Class<D> dataClass) {
        this.dataClass = dataClass;
    }

    @Override
    public void handle(int rowIndex, D data) {
        if (data != null) {
            list.add(data);
        }
    }

    /**
     * 获取所有数据
     *
     * @return 所有数据
     */
    public List<D> getResult() {
        return list;
    }

    @Override
    public Class<D> getDataClass() {
        return dataClass;
    }

}
