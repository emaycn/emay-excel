package cn.emay.excel.write.data;

import cn.emay.excel.common.schema.base.SheetSchema;

/**
 * 基于定义的表数据获取器
 * 
 * @author Frank
 *
 * @param <D>
 */
public interface SchemaSheetDataGetter<D> extends SheetDataGetter<D> {

	/**
	 * 表定义
	 * 
	 * @return
	 */
	public SheetSchema getSheetSchema();

}
