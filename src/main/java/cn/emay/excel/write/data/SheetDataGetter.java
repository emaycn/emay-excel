package cn.emay.excel.write.data;

/**
 * 
 * 基于Excel定义的数据处理器
 * 
 * @author Frank
 *
 * @param <D>
 */
public interface SheetDataGetter<D> {

	/**
	 * 获取数据
	 * 
	 * @param rowIndex
	 *            行号[从0开始]
	 * 
	 * @return
	 */
	D getData(int rowIndex);

	/**
	 * 是否有数据
	 * 
	 * @param rowIndex
	 *            行号[从0开始]
	 * 
	 * @return
	 */
	boolean hasData(int rowIndex);

	Class<D> getDataClass();

}
