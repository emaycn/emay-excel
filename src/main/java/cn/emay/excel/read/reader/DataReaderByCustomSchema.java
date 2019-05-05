package cn.emay.excel.read.reader;

import cn.emay.excel.schema.base.SheetSchema;

public interface DataReaderByCustomSchema<D> extends DataReader<D> {

	SheetSchema<D> getCustomSheetSchema();

}
