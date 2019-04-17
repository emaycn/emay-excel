package cn.emay.excel.read.handler.relize;

/**
 * 
 * 基于Schema读取到的数据处理器
 * 
 * @author Frank
 *
 * @param <D>
 */
public interface DataReader<D> {

	/**
	 * 处理一行数据
	 * 
	 * @param data
	 *            数据
	 * @param rowIndex
	 *            行号
	 */
	void handlerRowData(int rowIndex, D data);

	/**
	 * 获取Shema定义
	 * 
	 * @return
	 */
	Class<D> getShemaClass();

}
