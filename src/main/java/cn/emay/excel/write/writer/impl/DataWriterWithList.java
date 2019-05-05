package cn.emay.excel.write.writer.impl;

import java.util.List;

import cn.emay.excel.write.writer.DataWriter;

public class DataWriterWithList<D> implements DataWriter<D> {

	private List<D> datas;
	private int size;
	private Class<D> dataClass;

	public DataWriterWithList(List<D> datas, Class<D> dataClass) {
		this.datas = datas;
		size = datas.size();
		this.dataClass = dataClass;
	}

	@Override
	public D getData(int rowIndex) {
		return datas.get(rowIndex);
	}

	@Override
	public boolean hasData(int rowIndex) {
		return rowIndex < size;
	}

	@Override
	public Class<D> getDataClass() {
		return dataClass;
	}

}
