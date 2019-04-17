package cn.emay.excel.writer.handler;

import java.util.List;

import cn.emay.excel.common.Person;
import cn.emay.excel.write.handler.impl.DataWriter;

/**
 * 
 * @author Frank
 *
 */
public class PersonDataWriter implements DataWriter<Person> {

	private List<Person> datas;
	int size;

	public PersonDataWriter(List<Person> datas) {
		this.datas = datas;
		size = datas.size();
	}

	@Override
	public Person getData(int rowIndex) {
		return datas.get(rowIndex);
	}

	@Override
	public boolean hasData(int rowIndex) {
		return rowIndex < size;
	}

	@Override
	public Class<Person> getShemaClass() {
		return Person.class;
	}

}
