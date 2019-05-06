package cn.emay.excel.read.handler;

/**
 * 
 * 基于Schema读取到的数据处理器
 * 
 * @author Frank
 *
 * @param <D>
 */
public interface SheetDataHandler<D> {

	/**
	 * 处理一行数据
	 * 
	 * @param data
	 *            数据
	 * @param rowIndex
	 *            行号
	 */
	void handle(int rowIndex, D data);

	/**
	 * 获取数据Class
	 * 
	 * @return
	 */
	Class<D> getDataClass();

}
