package cn.emay.excel.read.handler;

import org.apache.poi.ss.usermodel.Cell;

/**
 * 读处理器
 * 
 * @author Frank
 *
 */
public interface SheetReadHandler {

	/**
	 * 开始读取的行数[从0开始]
	 * 
	 * @return
	 */
	int getStartReadRowIndex();

	/**
	 * 读取到哪一行结束<br/>
	 * 如果<0 则全部读取；
	 * 
	 * @return
	 */
	int getEndReadRowIndex();

	/**
	 * 开始读取
	 * 
	 * @param sheetIndex
	 *            sheet 序号
	 * @param sheetName
	 *            sheet名字
	 */
	void begin(int sheetIndex, String sheetName);

	/**
	 * 开始读取新的一行
	 * 
	 * @param rowIndex
	 *            行号[从0开始]
	 */
	void beginRow(int rowIndex);

	/**
	 * 处理Cell
	 * 
	 * @param rowIndex
	 *            行号
	 * @param columnIndex
	 *            列号[从0开始]
	 * @param cell
	 *            列单元格
	 */
	void handleXlsCell(int rowIndex, int columnIndex, Cell cell);

	/**
	 * 处理Cell
	 * 
	 * @param rowIndex
	 *            行号
	 * @param columnIndex
	 *            列号[从0开始]
	 * @param formatIndex
	 *            列单元格数据类型
	 * @param value
	 *            列值
	 */
	void handleXlsxCell(int rowIndex, int columnIndex, int formatIndex, String value);

	/**
	 * 结束一行的读取
	 * 
	 * @param rowIndex
	 *            行号[从0开始]
	 */
	void endRow(int rowIndex);

	/**
	 * 结束读取流程
	 * 
	 * @param sheetIndex
	 *            sheet 序号
	 * @param sheetName
	 *            sheet名字
	 */
	void end(int sheetIndex, String sheetName);

}
