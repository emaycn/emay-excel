package cn.emay.excel.write.data;

/**
 * 写入数据
 *
 * @author Frank
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

    /**
     *
     */
    public WriteData() {

    }

    /**
     * @param sheetIndex  表序号
     * @param rowIndex    行号
     * @param columnIndex 列号
     * @param data        数据
     * @param express     表达式(写入日期时：格式化日期; 写入Double、BigDecimal时：是保留的小数点后数字个数)
     */
    public WriteData(int sheetIndex, int rowIndex, int columnIndex, Object data, String express) {
        this.sheetIndex = sheetIndex;
        this.rowIndex = rowIndex;
        this.columnIndex = columnIndex;
        this.data = data;
        this.express = express;
    }

    /**
     * 获取坐标
     *
     * @return 坐标
     */
    public int[] getCoordinate() {
        return new int[]{sheetIndex, rowIndex, columnIndex};
    }

    /**
     * 获取表序号
     *
     * @return 表序号
     */
    public int getSheetIndex() {
        return sheetIndex;
    }

    /**
     * 传入表序号
     *
     * @param sheetIndex 表序号
     */
    public void setSheetIndex(int sheetIndex) {
        this.sheetIndex = sheetIndex;
    }

    /**
     * 获取行号
     *
     * @return 行号
     */
    public int getRowIndex() {
        return rowIndex;
    }

    /**
     * 传入行号
     *
     * @param rowIndex 行号
     */
    public void setRowIndex(int rowIndex) {
        this.rowIndex = rowIndex;
    }

    /**
     * 获取列号
     *
     * @return 列号
     */
    public int getColumnIndex() {
        return columnIndex;
    }

    /**
     * 传入列号
     *
     * @param columnIndex 列号
     */
    public void setColumnIndex(int columnIndex) {
        this.columnIndex = columnIndex;
    }

    /**
     * 获取数据
     *
     * @return 数据
     */
    public Object getData() {
        return data;
    }

    /**
     * 传入数据
     *
     * @param data 数据
     */
    public void setData(Object data) {
        this.data = data;
    }

    /**
     * 获取表达式
     *
     * @return 表达式
     */
    public String getExpress() {
        return express;
    }

    /**
     * 传入表达式
     *
     * @param express 表达式
     */
    public void setExpress(String express) {
        this.express = express;
    }

}