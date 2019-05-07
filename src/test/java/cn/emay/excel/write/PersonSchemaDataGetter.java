package cn.emay.excel.write;

import java.util.List;

import cn.emay.excel.common.Person;
import cn.emay.excel.common.schema.base.SheetSchema;
import cn.emay.excel.write.data.SchemaSheetDataGetter;

/**
 * 
 * @author Frank
 *
 */
public class PersonSchemaDataGetter extends PersonDataGetter implements SchemaSheetDataGetter<Person> {

	public PersonSchemaDataGetter(List<Person> datas) {
		super(datas);
	}

	@Override
	public SheetSchema getSheetSchema() {
		return new SheetSchema(Person.class);
	}

}
