package cn.emay.excel.reader.handler;

import java.util.ArrayList;
import java.util.List;

import cn.emay.excel.common.Person;
import cn.emay.excel.read.handler.relize.DataReader;

public class PersonDataReader implements DataReader<Person> {
	
	private List<Person> list = new ArrayList<>();

	@Override
	public void handlerRowData(int rowIndex, Person data) {
		list.add(data);
	}

	@Override
	public Class<Person> getShemaClass() {
		return Person.class;
	}
	
	public List<Person> getDatas() {
		return list;
	}
}