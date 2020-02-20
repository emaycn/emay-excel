package cn.emay.excel.read.core;

import java.io.File;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

import cn.emay.excel.read.reader.SheetReader;

/**
 * 统一读取器
 * 
 * @author Frank
 *
 */
public abstract class BaseReader {

	/**
	 * 从输入流中读取Excel表格<br/>
	 * 
	 * @param is
	 *            输入流
	 * @param sheetIndex
	 *            Sheet Index
	 * @param handler
	 *            Sheet读取处理器
	 */
	public void readBySheetIndex(InputStream is, int sheetIndex, SheetReader handler) {
		if (handler == null) {
			throw new IllegalArgumentException("handler is null");
		}
		if (sheetIndex < 0) {
			throw new IllegalArgumentException("sheetIndex begin with 0 , and must bigger than -1");
		}
		Map<Integer, SheetReader> handlersByName = new HashMap<>(1);
		handlersByName.put(sheetIndex, handler);
		readBySheetIndexs(is, handlersByName);
	}

	/**
	 * 从输入流中读取Excel表格<br/>
	 * 
	 * @param is
	 *            输入流
	 * @param handlers
	 *            Excel表处理器(handlers顺序号即为读取ExccelSheet的编号)
	 */
	public void readByOrder(InputStream is, SheetReader... handlers) {
		if (handlers == null || handlers.length == 0) {
			throw new IllegalArgumentException("handlers is null");
		}
		Map<Integer, SheetReader> handlersByIndex = new HashMap<>(handlers.length);
		for (int i = 0; i < handlers.length; i++) {
			handlersByIndex.put(i, handlers[i]);
		}
		readBySheetIndexs(is, handlersByIndex);
	}

	/**
	 * 从输入流中读取Excel表格<br/>
	 * 按照表序号匹配读取处理器<br/>
	 * 
	 * @param is
	 *            输入流
	 * @param handlersByIndex
	 *            按照Index匹配的Sheet读取处理器集合
	 */
	public void readBySheetIndexs(InputStream is, Map<Integer, SheetReader> handlersByIndex) {
		if (handlersByIndex == null || handlersByIndex.size() == 0) {
			throw new IllegalArgumentException("handlers is null");
		}
		read(is, handlersByIndex, null);
	}

	/**
	 * 从输入流中读取Excel表格<br/>
	 * 
	 * @param is
	 *            输入流
	 * @param sheetName
	 *            Sheet页名字
	 * @param handler
	 *            Sheet读取处理器
	 */
	public void readBySheetName(InputStream is, String sheetName, SheetReader handler) {
		if (handler == null) {
			throw new IllegalArgumentException("handler is null");
		}
		if (sheetName == null) {
			throw new IllegalArgumentException("sheetName is null");
		}
		Map<String, SheetReader> handlersByName = new HashMap<>(1);
		handlersByName.put(sheetName, handler);
		readBySheetNames(is, handlersByName);
	}

	/**
	 * 从输入流中读取Excel表格<br/>
	 * 按照表名匹配读取处理器<br/>
	 * 
	 * @param is
	 *            输入流
	 * @param handlersByName
	 *            按照表名匹配的Sheet读取处理器集合
	 */
	public void readBySheetNames(InputStream is, Map<String, SheetReader> handlersByName) {
		if (handlersByName == null || handlersByName.size() == 0) {
			throw new IllegalArgumentException("handlers is null");
		}
		read(is, null, handlersByName);
	}

	/**
	 * 从输入流中读取Excel表格<br/>
	 * 所有处理器依次进行匹配
	 * 
	 * @param isByIndex
	 *            是否按照Index匹配读取处理器
	 * @param handlersByIndex
	 *            按照Index匹配的Sheet读取处理器集合
	 * @param handlersByName
	 *            按照表名匹配的Sheet读取处理器集合
	 */
	public abstract void read(InputStream is, Map<Integer, SheetReader> handlersByIndex, Map<String, SheetReader> handlersByName);

	/**
	 * 从输入流中读取Excel表格<br/>
	 * 
	 * @param is
	 *            输入流
	 * @param sheetIndex
	 *            Sheet Index
	 * @param handler
	 *            Sheet读取处理器
	 */
	public void readBySheetIndex(File file, int sheetIndex, SheetReader handler) {
		if (handler == null) {
			throw new IllegalArgumentException("handler is null");
		}
		if (sheetIndex < 0) {
			throw new IllegalArgumentException("sheetIndex begin with 0 , and must bigger than -1");
		}
		Map<Integer, SheetReader> handlersByName = new HashMap<>(1);
		handlersByName.put(sheetIndex, handler);
		readBySheetIndexs(file, handlersByName);
	}

	/**
	 * 从输入流中读取Excel表格<br/>
	 * 
	 * @param is
	 *            输入流
	 * @param handlers
	 *            Excel表处理器(handlers顺序号即为读取ExccelSheet的编号)
	 */
	public void readByOrder(File file, SheetReader... handlers) {
		if (handlers == null || handlers.length == 0) {
			throw new IllegalArgumentException("handlers is null");
		}
		Map<Integer, SheetReader> handlersByIndex = new HashMap<>(handlers.length);
		for (int i = 0; i < handlers.length; i++) {
			handlersByIndex.put(i, handlers[i]);
		}
		readBySheetIndexs(file, handlersByIndex);
	}

	/**
	 * 从输入流中读取Excel表格<br/>
	 * 按照表序号匹配读取处理器<br/>
	 * 
	 * @param is
	 *            输入流
	 * @param handlersByIndex
	 *            按照Index匹配的Sheet读取处理器集合
	 */
	public void readBySheetIndexs(File file, Map<Integer, SheetReader> handlersByIndex) {
		if (handlersByIndex == null || handlersByIndex.size() == 0) {
			throw new IllegalArgumentException("handlers is null");
		}
		read(file, handlersByIndex, null);
	}

	/**
	 * 从输入流中读取Excel表格<br/>
	 * 
	 * @param is
	 *            输入流
	 * @param sheetName
	 *            Sheet页名字
	 * @param handler
	 *            Sheet读取处理器
	 */
	public void readBySheetName(File file, String sheetName, SheetReader handler) {
		if (handler == null) {
			throw new IllegalArgumentException("handler is null");
		}
		if (sheetName == null) {
			throw new IllegalArgumentException("sheetName is null");
		}
		Map<String, SheetReader> handlersByName = new HashMap<>(1);
		handlersByName.put(sheetName, handler);
		readBySheetNames(file, handlersByName);
	}

	/**
	 * 从输入流中读取Excel表格<br/>
	 * 按照表名匹配读取处理器<br/>
	 * 
	 * @param is
	 *            输入流
	 * @param handlersByName
	 *            按照表名匹配的Sheet读取处理器集合
	 */
	public void readBySheetNames(File file, Map<String, SheetReader> handlersByName) {
		if (handlersByName == null || handlersByName.size() == 0) {
			throw new IllegalArgumentException("handlers is null");
		}
		read(file, null, handlersByName);
	}

	/**
	 * 从输入流中读取Excel表格<br/>
	 * 所有处理器依次进行匹配
	 * 
	 * @param isByIndex
	 *            是否按照Index匹配读取处理器
	 * @param handlersByIndex
	 *            按照Index匹配的Sheet读取处理器集合
	 * @param handlersByName
	 *            按照表名匹配的Sheet读取处理器集合
	 */
	public abstract void read(File file, Map<Integer, SheetReader> handlersByIndex, Map<String, SheetReader> handlersByName);

}
