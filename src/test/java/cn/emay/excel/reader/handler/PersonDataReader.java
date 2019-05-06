package cn.emay.excel.reader.handler;

import java.util.ArrayList;
import java.util.List;

import cn.emay.excel.common.Person;
import cn.emay.excel.read.handler.SheetDataHandler;

/**
 * 
 * @author Frank
 *
 */
public class PersonDataReader implements SheetDataHandler<Person> {

	private List<Person> list = new ArrayList<>();

	@Override
	public void handle(int rowIndex, Person data) {
		list.add(data);
	}

	public List<Person> getDatas() {
		return list;
	}

	@Override
	public Class<Person> getDataClass() {
		return Person.class;
	}
}