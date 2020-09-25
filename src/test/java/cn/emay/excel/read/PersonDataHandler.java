package cn.emay.excel.read;

import cn.emay.excel.common.Person;
import cn.emay.excel.read.handler.SheetDataHandler;

import java.util.ArrayList;
import java.util.List;

/**
 * @author Frank
 */
public class PersonDataHandler implements SheetDataHandler<Person> {

    private final List<Person> list = new ArrayList<>();

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