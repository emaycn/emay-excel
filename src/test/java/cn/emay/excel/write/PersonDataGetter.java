package cn.emay.excel.write;

import java.util.List;

import cn.emay.excel.common.Person;
import cn.emay.excel.write.data.SheetDataGetter;

/**
 * 
 * @author Frank
 *
 */
public class PersonDataGetter implements SheetDataGetter<Person> {

	private List<Person> datas;
	int size;

	public PersonDataGetter(List<Person> datas) {
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
	public Class<Person> getDataClass() {
		return Person.class;
	}

}
