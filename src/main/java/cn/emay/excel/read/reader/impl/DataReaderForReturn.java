package cn.emay.excel.read.reader.impl;

import java.util.ArrayList;
import java.util.List;

import cn.emay.excel.read.reader.DataReader;

/**
 * 结果全部保存到内存中统一返回的数据处理器
 * 
 * @author Frank
 *
 * @param <D>
 */
public class DataReaderForReturn<D> implements DataReader<D> {

	/**
	 * 所有数据
	 */
	private List<D> list = new ArrayList<>();
	
	private Class<D> dataClass;
	
	public DataReaderForReturn(Class<D> dataClass) {
		this.dataClass = dataClass;
	}

	@Override
	public void handlerRowData(int rowIndex, D data) {
		if (data != null) {
			list.add(data);
		}
	}

	/**
	 * 获取所有数据
	 * 
	 * @return
	 */
	public List<D> getResult() {
		return list;
	}

	@Override
	public Class<D> getDataClass() {
		return dataClass;
	}

}
