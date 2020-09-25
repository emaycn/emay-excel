package cn.emay.excel.read.handler;

/**
 * 基于Schema的表数据处理器
 *
 * @param <D>
 * @author Frank
 */
public interface SheetDataHandler<D> {

    /**
     * 处理一行数据
     *
     * @param data     数据
     * @param rowIndex 行号
     */
    void handle(int rowIndex, D data);

    /**
     * 获取数据Class
     *
     * @return 数据Class
     */
    Class<D> getDataClass();

}
