package cn.emay.excel.write;

import cn.emay.excel.common.Person;
import cn.emay.excel.write.data.SheetDataGetter;

import java.util.List;

/**
 * @author Frank
 */
public class PersonDataGetter implements SheetDataGetter<Person> {

    private final List<Person> datas;
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
