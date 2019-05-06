package cn.emay.excel.read.handler;

import cn.emay.excel.common.schema.base.SheetSchema;

public interface SchemaSheetDataHandler<D> extends SheetDataHandler<D> {

	public SheetSchema<D> getSheetSchema();

}
