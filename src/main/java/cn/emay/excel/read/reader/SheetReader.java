package cn.emay.excel.read.reader;

import org.apache.poi.ss.usermodel.Cell;

/**
 * 读处理器
 *
 * @author Frank
 */
public interface SheetReader {

    /**
     * 开始读取的行数[从0开始]
     *
     * @return 开始读取的行数[从0开始]
     */
    int getStartReadRowIndex();

    /**
     * 读取到哪一行结束<br/>
     * 如果<0 则全部读取；
     *
     * @return 读取到哪一行结束
     */
    int getEndReadRowIndex();

    /**
     * 开始读取
     *
     * @param sheetIndex sheet 序号
     * @param sheetName  sheet名字
     */
    void begin(int sheetIndex, String sheetName);

    /**
     * 开始读取新的一行
     *
     * @param rowIndex 行号[从0开始]
     */
    void beginRow(int rowIndex);

    /**
     * 处理Xls的单元格<br/>
     * 解析98/03版本.xls后缀的Excel文件是调用此方法
     *
     * @param rowIndex    行号
     * @param columnIndex 列号[从0开始]
     * @param cell        列单元格
     */
    void handleXlsCell(int rowIndex, int columnIndex, Cell cell);

    /**
     * 处理Xlsx的单元格<br/>
     * 解析07+版本.xlsx后缀的Excel文件是调用此方法
     *
     * @param rowIndex    行号
     * @param columnIndex 列号[从0开始]
     * @param value       列值
     */
    void handleXlsxCell(int rowIndex, int columnIndex, String value);

    /**
     * 结束一行的读取
     *
     * @param rowIndex 行号[从0开始]
     */
    void endRow(int rowIndex);

    /**
     * 结束读取流程
     *
     * @param sheetIndex sheet 序号
     * @param sheetName  sheet名字
     */
    void end(int sheetIndex, String sheetName);

}
