package cn.emay.excel.write.data;

import cn.emay.excel.common.schema.base.SheetSchema;

public interface SchemaSheetDataGetter<D> extends SheetDataGetter<D>{

	public SheetSchema<D> getSheetSchema();

}
