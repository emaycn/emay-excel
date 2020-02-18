package cn.emay.excel.write.writer;

import org.apache.poi.ss.usermodel.Cell;

/**
 * Excel Sheet写入处理器
 * 
 * @author Frank
 *
 */
public interface SheetWriter {

	/**
	 * 写入的表名
	 * 
	 * @return
	 */
	String getSheetName();

	/**
	 * 最大的列数[从0开始]
	 * 
	 * @return
	 */
	int getMaxColumnIndex();

	/**
	 * 是否自适应宽度
	 * 
	 * @return
	 */
	boolean isAutoWidth();

	/**
	 * 开始写入
	 * 
	 * @param sheetIndex sheet编号[从0开始]
	 */
	void begin(int sheetIndex);

	/**
	 * 是否有数据
	 * 
	 * @param rowIndex 行号[从0开始]
	 * 
	 * @return
	 */
	boolean hasRow(int rowIndex);

	/**
	 * 开始写入新的一行
	 * 
	 * @param rowIndex 行号[从0开始]
	 */
	void beginRow(int rowIndex);

	/**
	 * 写入数据
	 * 
	 * @param cell        单元格
	 * @param rowIndex    行号
	 * @param columnIndex 列号
	 */
	void writeCell(Cell cell, int rowIndex, int columnIndex);

	/**
	 * 结束一行的写入
	 * 
	 * @param rowIndex 行号[从0开始]
	 */
	void endRow(int rowIndex);

	/**
	 * 结束写入流程
	 * 
	 * @param sheetIndex sheet编号[从0开始]
	 */
	void end(int sheetIndex);

}
