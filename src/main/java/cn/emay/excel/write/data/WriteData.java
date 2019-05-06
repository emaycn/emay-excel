package cn.emay.excel.write.data;

/**
 * 写入数据
 * 
 * @author Frank
 *
 */
public class WriteData {

	/**
	 * 表序号
	 */
	private int sheetIndex;

	/**
	 * 行号
	 */
	private int rowIndex;

	/**
	 * 列号
	 */
	private int columnIndex;

	/**
	 * 数据
	 */
	private Object data;

	/**
	 * 数据格式表达式
	 */
	private String express;

	public WriteData() {

	}

	/**
	 * 
	 * @param sheetIndex
	 *            表序号
	 * @param rowIndex
	 *            行号
	 * @param columnIndex
	 *            列号
	 * @param data
	 *            数据
	 */
	public WriteData(int sheetIndex, int rowIndex, int columnIndex, Object data, String express) {
		this.sheetIndex = sheetIndex;
		this.rowIndex = rowIndex;
		this.columnIndex = columnIndex;
		this.data = data;
		this.express = express;
	}

	public int getSheetIndex() {
		return sheetIndex;
	}

	public void setSheetIndex(int sheetIndex) {
		this.sheetIndex = sheetIndex;
	}

	public int getRowIndex() {
		return rowIndex;
	}

	public void setRowIndex(int rowIndex) {
		this.rowIndex = rowIndex;
	}

	public int getColumnIndex() {
		return columnIndex;
	}

	public void setColumnIndex(int columnIndex) {
		this.columnIndex = columnIndex;
	}

	public Object getData() {
		return data;
	}

	public void setData(Object data) {
		this.data = data;
	}

	public String getExpress() {
		return express;
	}

	public void setExpress(String express) {
		this.express = express;
	}

}