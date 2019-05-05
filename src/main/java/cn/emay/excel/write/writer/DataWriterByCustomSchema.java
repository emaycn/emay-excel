package cn.emay.excel.write.writer;

import cn.emay.excel.schema.base.SheetSchema;

public interface DataWriterByCustomSchema<D> extends DataWriter<D>{
	
	SheetSchema<D> getSheetSchema();

}
