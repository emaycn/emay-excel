package cn.emay.excel.read;

import cn.emay.excel.common.Person;
import cn.emay.excel.common.schema.base.SheetSchema;
import cn.emay.excel.read.handler.SchemaSheetDataHandler;

/**
 * 
 * @author Frank
 *
 */
public class PersonSchemaDataHandler extends PersonDataHandler implements SchemaSheetDataHandler<Person>{

	@Override
	public SheetSchema getSheetSchema() {
		return new SheetSchema(Person.class);
	}

}
