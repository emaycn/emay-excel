package cn.emay.excel.write.data;

/**
 * 表数据获取器
 *
 * @param <D>
 * @author Frank
 */
public interface SheetDataGetter<D> {

    /**
     * 获取数据
     *
     * @param rowIndex 行号[从0开始]
     * @return 数据
     */
    D getData(int rowIndex);

    /**
     * 是否有数据
     *
     * @param rowIndex 行号[从0开始]
     * @return 是否有数据
     */
    boolean hasData(int rowIndex);

    /**
     * 数据Class
     *
     * @return 数据Class
     */
    Class<D> getDataClass();

}
