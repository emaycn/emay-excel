package cn.emay.excel.read.core;

import cn.emay.excel.read.reader.SheetReader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * XLS读取器
 *
 * @author Frank
 */
public class XlsReader extends BaseReader {

    @Override
    public void read(InputStream is, Map<Integer, SheetReader> handlersByIndex, Map<String, SheetReader> handlersByName) {
        if (is == null) {
            throw new IllegalArgumentException("InputStream is null");
        }
        try (
                InputStream isNew = is;
                Workbook workbook = new HSSFWorkbook(isNew)
        ) {
            readWorkbook(workbook, handlersByIndex, handlersByName);
        } catch (IOException e) {
            throw new IllegalArgumentException(e);
        }
    }

    /**
     * 从Workbook中读取Excel表格<br/>
     *
     * @param workbook   workbook
     * @param sheetIndex Sheet Index
     * @param handler    Sheet读取处理器
     */
    public void readBySheetIndex(Workbook workbook, int sheetIndex, SheetReader handler) {
        if (handler == null) {
            throw new IllegalArgumentException("handler is null");
        }
        if (sheetIndex < 0) {
            throw new IllegalArgumentException("sheetIndex begin with 0 , and must bigger than 0");
        }
        Map<Integer, SheetReader> handlers = new HashMap<>(1);
        handlers.put(sheetIndex, handler);
        readBySheetIndexs(workbook, handlers);
    }

    /**
     * 从Workbook中读取Excel表格<br/>
     *
     * @param workbook Workbook
     * @param handlers Excel表处理器(handlers顺序号即为读取ExccelSheet的编号)
     */
    public void readByOrder(Workbook workbook, SheetReader... handlers) {
        if (handlers == null || handlers.length == 0) {
            throw new IllegalArgumentException("handlers is null");
        }
        Map<Integer, SheetReader> handlersByIndex = new HashMap<>(handlers.length);
        for (int i = 0; i < handlers.length; i++) {
            handlersByIndex.put(i, handlers[i]);
        }
        readBySheetIndexs(workbook, handlersByIndex);
    }

    /**
     * 从Workbook中读取Excel表格<br/>
     * 按照表序号匹配读取处理器<br/>
     *
     * @param workbook        workbook
     * @param handlersByIndex 按照Index匹配的Sheet读取处理器集合
     */
    public void readBySheetIndexs(Workbook workbook, Map<Integer, SheetReader> handlersByIndex) {
        if (handlersByIndex == null || handlersByIndex.size() == 0) {
            throw new IllegalArgumentException("handlers is null");
        }
        readWorkbook(workbook, handlersByIndex, null);
    }

    /**
     * 从Workbook中读取Excel表格<br/>
     *
     * @param workbook  workbook
     * @param sheetName Sheet页名字
     * @param handler   Sheet读取处理器
     */
    public void readBySheetName(Workbook workbook, String sheetName, SheetReader handler) {
        if (handler == null) {
            throw new IllegalArgumentException("handler is null");
        }
        if (sheetName == null) {
            throw new IllegalArgumentException("sheetName is null");
        }
        Map<String, SheetReader> handlersByName = new HashMap<>(1);
        handlersByName.put(sheetName, handler);
        readBySheetNames(workbook, handlersByName);
    }

    /**
     * 从Workbook中读取Excel表格<br/>
     * 按照表名匹配读取处理器<br/>
     *
     * @param workbook       workbook
     * @param handlersByName 按照表名匹配的Sheet读取处理器集合
     */
    public void readBySheetNames(Workbook workbook, Map<String, SheetReader> handlersByName) {
        if (handlersByName == null || handlersByName.size() == 0) {
            throw new IllegalArgumentException("handlers is null");
        }
        readWorkbook(workbook, null, handlersByName);
    }

    /**
     * 从Workbook中读取Excel表格<br/>
     * 所有处理器依次进行匹配
     *
     * @param workbook        workbook
     * @param handlersByIndex 按照Index匹配的Sheet读取处理器集合
     * @param handlersByName  按照表名匹配的Sheet读取处理器集合
     */
    public void readWorkbook(Workbook workbook, Map<Integer, SheetReader> handlersByIndex, Map<String, SheetReader> handlersByName) {
        if (workbook == null) {
            throw new IllegalArgumentException("workbook is null");
        }
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            if (sheet == null) {
                continue;
            }
            String name = sheet.getSheetName();
            SheetReader readHander = null;
            if (handlersByIndex != null) {
                readHander = handlersByIndex.get(i);
            }
            if (readHander == null && handlersByName != null) {
                readHander = handlersByName.get(name);
            }
            if (readHander == null) {
                continue;
            }
            readSheet(sheet, readHander);
        }
    }

    /**
     * 读取Sheet
     *
     * @param sheet   sheet页
     * @param handler Sheet读取处理器
     */
    public void readSheet(Sheet sheet, SheetReader handler) {
        if (sheet == null) {
            throw new IllegalArgumentException("sheet is null");
        }
        if (handler == null) {
            throw new IllegalArgumentException("handlers is null");
        }
        int index = sheet.getWorkbook().getSheetIndex(sheet);
        String name = sheet.getSheetName();
        handler.begin(index, name);
        int startReadRowIndex = handler.getStartReadRowIndex();
        int endReadRowIndex = handler.getEndReadRowIndex();
        int begin = Math.max(startReadRowIndex, 0);
        for (int j = begin; j <= sheet.getLastRowNum(); j++) {
            if (endReadRowIndex >= 0 && j > endReadRowIndex) {
                break;
            }
            Row row = sheet.getRow(j);
            if (row == null) {
                continue;
            }
            handler.beginRow(j);
            for (int k = 0; k <= row.getLastCellNum(); k++) {
                Cell cell = row.getCell(k);
                handler.handleXlsCell(j, k, cell);
            }
            handler.endRow(j);
        }
        handler.end(index, name);
    }

    @Override
    public void read(File file, Map<Integer, SheetReader> handlersByIndex, Map<String, SheetReader> handlersByName) {
        if (file == null) {
            throw new IllegalArgumentException("File is null");
        }
        try (
                FileInputStream is = new FileInputStream(file);
                Workbook workbook = new HSSFWorkbook(is)
        ) {
            readWorkbook(workbook, handlersByIndex, handlersByName);
        } catch (IOException e) {
            throw new IllegalArgumentException(e);
        }
    }

}