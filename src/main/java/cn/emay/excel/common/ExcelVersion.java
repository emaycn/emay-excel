package cn.emay.excel.common;

/**
 * Excel版本
 *
 * @author Frank
 */
public enum ExcelVersion {

    /**
     * 95-03版本，.xls后缀
     */
    XLS(".xls"),
    /**
     * 07+版本，.xlsx后缀
     */
    XLSX(".xlsx");

    private String suffix;

    ExcelVersion(String suffix) {
        this.setSuffix(suffix);
    }

    /**
     * 获取后缀
     *
     * @return 后缀
     */
    public String getSuffix() {
        return suffix;
    }

    public void setSuffix(String suffix) {
        this.suffix = suffix;
    }

}
