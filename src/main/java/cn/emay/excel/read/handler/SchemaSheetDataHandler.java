package cn.emay.excel.read.handler;

import cn.emay.excel.common.schema.base.SheetSchema;

/**
 * 基于定义的表格数据处理器
 * 
 * @author Frank
 *
 * @param <D>
 */
public interface SchemaSheetDataHandler<D> extends SheetDataHandler<D> {

	/**
	 * 获取定义
	 * 
	 * @return
	 */
	public SheetSchema<D> getSheetSchema();

}
