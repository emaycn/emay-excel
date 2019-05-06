package cn.emay.excel.write.data;

/**
 * 
 * 表数据获取器
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

	/**
	 * 数据Class
	 * 
	 * @return
	 */
	Class<D> getDataClass();

}
