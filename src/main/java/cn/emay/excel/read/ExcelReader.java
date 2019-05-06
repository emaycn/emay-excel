package cn.emay.excel.read;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import cn.emay.excel.common.ExcelVersion;
import cn.emay.excel.common.schema.base.SheetSchema;
import cn.emay.excel.read.core.CoreReader;
import cn.emay.excel.read.core.XlsReader;
import cn.emay.excel.read.core.XlsxReader;
import cn.emay.excel.read.handler.SchemaSheetDataHandler;
import cn.emay.excel.read.handler.SheetDataHandler;
import cn.emay.excel.read.reader.SheetReader;
import cn.emay.excel.read.reader.impl.SchemaSheetReader;

/**
 * Excel基础读取<br/>
 * XLSX统一采用SAX方式读取
 * 
 * @author Frank
 *
 */
public class ExcelReader {

	/**
	 * 从Excel文件中读取第一个表格<br/>
	 * dataClass实现了@ExcelSheet注解,其字段实现了@ExcelColumn注解
	 * 
	 * @param excelPath
	 *            路径
	 * @param dataClass
	 *            数据Class
	 * @return 数据
	 */
	public static <D> List<D> readFirstSheet(String excelPath, Class<D> dataClass) {
		return readBySheetIndex(excelPath, 0, dataClass);
	}

	/**
	 * 从Excel文件中读取一个表格<br/>
	 * dataClass实现了@ExcelSheet注解,其字段实现了@ExcelColumn注解
	 * 
	 * @param excelPath
	 *            路径
	 * @param sheetName
	 *            Sheet页名字
	 * @param dataClass
	 *            数据Class
	 * @return 数据
	 */
	public static <D> List<D> readBySheetName(String excelPath, String sheetName, Class<D> dataClass) {
		ReturnSchemaDataReader<D> dataReader = new ReturnSchemaDataReader<D>(dataClass);
		readBySheetName(excelPath, sheetName, dataReader);
		return dataReader.getResult();
	}

	/**
	 * 从Excel文件中读取一个表格<br/>
	 * dataClass实现了@ExcelSheet注解,其字段实现了@ExcelColumn注解
	 * 
	 * @param excelPath
	 *            路径
	 * @param sheetIndex
	 *            Sheet Index
	 * @param dataClass
	 *            数据Class
	 * @return 数据
	 */
	public static <D> List<D> readBySheetIndex(String excelPath, int sheetIndex, Class<D> dataClass) {
		ReturnSchemaDataReader<D> dataReader = new ReturnSchemaDataReader<D>(dataClass);
		readBySheetIndex(excelPath, sheetIndex, dataReader);
		return dataReader.getResult();
	}

	/**
	 * 从Excel文件中读取第一个表格<br/>
	 * dataClass 实现了@ExcelSheet注解,其字段实现了@ExcelColumn注解
	 * 
	 * @param excelPath
	 *            路径
	 * @param dataHandler
	 *            数据处理器
	 * @return
	 */
	public static <D> void readFirstSheet(String excelPath, SheetDataHandler<D> dataHandler) {
		readBySheetIndex(excelPath, 0, dataHandler);
	}

	/**
	 * 从Excel文件中读取一个表格<br/>
	 * dataClass实现了@ExcelSheet注解,其字段实现了@ExcelColumn注解
	 * 
	 * @param excelPath
	 *            路径
	 * @param sheetName
	 *            Sheet页名字
	 * @param dataHandler
	 *            数据处理器
	 */
	public static <D> void readBySheetName(String excelPath, String sheetName, SheetDataHandler<D> dataHandler) {
		SchemaSheetReader<D> handler = new SchemaSheetReader<D>(new SheetSchema<D>(dataHandler.getDataClass()), dataHandler);
		readBySheetName(excelPath, sheetName, handler);
	}

	/**
	 * 从Excel文件中读取一个表格<br/>
	 * dataClass实现了@ExcelSheet注解,其字段实现了@ExcelColumn注解
	 * 
	 * @param excelPath
	 *            路径
	 * @param sheetIndex
	 *            Sheet Index
	 * @param dataHandler
	 *            数据处理器
	 */
	public static <D> void readBySheetIndex(String excelPath, int sheetIndex, SheetDataHandler<D> dataHandler) {
		SchemaSheetReader<D> handler = new SchemaSheetReader<D>(new SheetSchema<D>(dataHandler.getDataClass()), dataHandler);
		readBySheetIndex(excelPath, sheetIndex, handler);
	}

	/**
	 * 从Excel文件中读取第一个表格<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param schemaSheetDataHandler
	 *            表格定义读取器
	 * @return
	 */
	public static <D> void readFirstSheet(String excelPath, SchemaSheetDataHandler<D> schemaSheetDataHandler) {
		readBySheetIndex(excelPath, 0, schemaSheetDataHandler);
	}

	/**
	 * 从Excel文件中读取一个表格<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param sheetName
	 *            Sheet页名字
	 * @param schemaSheetDataHandler
	 *            表格定义读取器
	 */
	public static <D> void readBySheetName(String excelPath, String sheetName, SchemaSheetDataHandler<D> schemaSheetDataHandler) {
		SchemaSheetReader<D> handler = new SchemaSheetReader<D>(schemaSheetDataHandler.getSheetSchema(), schemaSheetDataHandler);
		readBySheetName(excelPath, sheetName, handler);
	}

	/**
	 * 从Excel文件中读取一个表格<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param sheetIndex
	 *            Sheet Index
	 * @param schemaSheetDataHandler
	 *            表格定义读取器
	 */
	public static <D> void readBySheetIndex(String excelPath, int sheetIndex, SchemaSheetDataHandler<D> schemaSheetDataHandler) {
		SchemaSheetReader<D> handler = new SchemaSheetReader<D>(schemaSheetDataHandler.getSheetSchema(), schemaSheetDataHandler);
		readBySheetIndex(excelPath, sheetIndex, handler);
	}

	/**
	 * 从文件中读取Excel表格<br/>
	 * 按照表序号匹配读取处理器<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param handlersByIndex
	 *            按照Index匹配的Sheet读取表格定义读取器
	 */
	@SuppressWarnings({ "rawtypes", "unchecked" })
	public static void readBySheetIndexsWithSchema(String excelPath, Map<Integer, SchemaSheetDataHandler> handlersByIndex) {
		if (handlersByIndex == null) {
			return;
		}
		Map<Integer, SheetReader> readers = new HashMap<Integer, SheetReader>();
		for (Integer index : handlersByIndex.keySet()) {
			SchemaSheetDataHandler<?> schemaSheetDataHandler = handlersByIndex.get(index);
			SchemaSheetReader<?> handler = new SchemaSheetReader(schemaSheetDataHandler.getSheetSchema(), schemaSheetDataHandler);
			readers.put(index, handler);
		}
		readBySheetIndexs(excelPath, readers);
	}

	/**
	 * 从文件中读取Excel表格<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param handlers
	 *            Excel表格定义读取器(handlers顺序号即为读取ExccelSheet的编号)
	 */
	@SuppressWarnings({ "rawtypes", "unchecked" })
	public static void readByOrder(String excelPath, SchemaSheetDataHandler... handlers) {
		if (handlers == null) {
			return;
		}
		SheetReader[] readers = new SheetReader[handlers.length];
		for (int i = 0; i < handlers.length; i++) {
			readers[i] = new SchemaSheetReader(handlers[i].getSheetSchema(), handlers[i]);
		}
		readByOrder(excelPath, readers);
	}

	/**
	 * 从文件中读取Excel表格<br/>
	 * 按照表名匹配读取处理器<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param handlersByName
	 *            按照表名匹配的Sheet读取表格定义读取器
	 */
	@SuppressWarnings({ "rawtypes", "unchecked" })
	public static void readBySheetNamesWithSchema(String excelPath, Map<String, SchemaSheetDataHandler> handlersByName) {
		if (handlersByName == null) {
			return;
		}
		Map<String, SheetReader> readers = new HashMap<String, SheetReader>();
		for (String name : handlersByName.keySet()) {
			SchemaSheetDataHandler<?> schemaSheetDataHandler = handlersByName.get(name);
			SchemaSheetReader<?> handler = new SchemaSheetReader(schemaSheetDataHandler.getSheetSchema(), schemaSheetDataHandler);
			readers.put(name, handler);
		}
		readBySheetNames(excelPath, readers);
	}

	/**
	 * 从文件中读取Excel表格第一个sheet页<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param reader
	 *            Sheet读取处理器[处理器实例不可复用]
	 */
	public static void readFirstSheet(String excelPath, SheetReader reader) {
		readBySheetIndex(excelPath, 0, reader);
	}

	/**
	 * 从文件中读取Excel表格<br/>
	 * 按照表格序号读取<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param sheetIndex
	 *            Sheet Index
	 * @param reader
	 *            Sheet读取处理器[处理器实例不可复用]
	 */
	public static void readBySheetIndex(String excelPath, int sheetIndex, SheetReader reader) {
		PathParser parser = new PathParser(excelPath);
		readBySheetIndex(parser.getInputStream(), parser.getVersion(), sheetIndex, reader);
	}

	/**
	 * 从文件中读取Excel表格<br/>
	 * 按照表序号匹配读取处理器<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param readersByIndex
	 *            按照Index匹配的Sheet读取处理器集合[处理器实例不可复用]
	 */
	public static void readBySheetIndexs(String excelPath, Map<Integer, SheetReader> readersByIndex) {
		PathParser parser = new PathParser(excelPath);
		readBySheetIndexs(parser.getInputStream(), parser.getVersion(), readersByIndex);
	}

	/**
	 * 从文件中读取Excel表格<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param readers
	 *            Excel表处理器(handlers顺序号即为读取ExccelSheet的编号)
	 */
	public static void readByOrder(String excelPath, SheetReader... readers) {
		PathParser parser = new PathParser(excelPath);
		readByOrder(parser.getInputStream(), parser.getVersion(), readers);
	}

	/**
	 * 从文件中读取Excel表格<br/>
	 * 按照表格名字读取<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param sheetName
	 *            Sheet页名字
	 * @param reader
	 *            Sheet读取处理器[处理器实例不可复用]
	 */
	public static void readBySheetName(String excelPath, String sheetName, SheetReader reader) {
		PathParser parser = new PathParser(excelPath);
		readBySheetName(parser.getInputStream(), parser.getVersion(), sheetName, reader);
	}

	/**
	 * 从文件中读取Excel表格<br/>
	 * 按照表名匹配读取处理器<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param readersByName
	 *            按照表名匹配的Sheet读取处理器集合[处理器实例不可复用]
	 */
	public static void readBySheetNames(String excelPath, Map<String, SheetReader> readersByName) {
		PathParser parser = new PathParser(excelPath);
		readBySheetNames(parser.getInputStream(), parser.getVersion(), readersByName);

	}

	/**
	 * 从文件中读取Excel表格第一个sheet页<br/>
	 * 
	 * @param is
	 *            输入流
	 * @param version
	 *            版本
	 * @param reader
	 *            Sheet读取处理器[处理器实例不可复用]
	 */
	public static void readFirstSheet(InputStream is, ExcelVersion version, SheetReader reader) {
		readBySheetIndex(is, version, 0, reader);
	}

	/**
	 * 从输入流中读取Excel表格<br/>
	 * 按照表格序号读取<br/>
	 * 
	 * @param is
	 *            输入流
	 * @param version
	 *            版本
	 * @param sheetIndex
	 *            Sheet Index
	 * @param reader
	 *            Sheet读取处理器[处理器实例不可复用]
	 */
	public static void readBySheetIndex(InputStream is, ExcelVersion version, int sheetIndex, SheetReader reader) {
		getReader(version).readBySheetIndex(is, sheetIndex, reader);
	}

	/**
	 * 从输入流中读取Excel表格<br/>
	 * 
	 * @param is
	 *            输入流
	 * @param readers
	 *            Excel表处理器(handlers顺序号即为读取ExccelSheet的编号)
	 */
	public static void readByOrder(InputStream is, ExcelVersion version, SheetReader... readers) {
		getReader(version).readByOrder(is, readers);
	}

	/**
	 * 从输入流中读取Excel表格<br/>
	 * 按照表序号匹配读取处理器<br/>
	 * 
	 * @param is
	 *            输入流
	 * @param version
	 *            版本
	 * @param readersByIndex
	 *            按照Index匹配的Sheet读取处理器集合[处理器实例不可复用]
	 */
	public static void readBySheetIndexs(InputStream is, ExcelVersion version, Map<Integer, SheetReader> readersByIndex) {
		getReader(version).readBySheetIndexs(is, readersByIndex);
	}

	/**
	 * 从输入流中读取Excel表格<br/>
	 * 按照表格名字读取<br/>
	 * 
	 * @param is
	 *            输入流
	 * @param version
	 *            版本
	 * @param sheetName
	 *            Sheet页名字
	 * @param reader
	 *            Sheet读取处理器[处理器实例不可复用]
	 */
	public static void readBySheetName(InputStream is, ExcelVersion version, String sheetName, SheetReader reader) {
		getReader(version).readBySheetName(is, sheetName, reader);
	}

	/**
	 * 从输入流中读取Excel表格<br/>
	 * 按照表名匹配读取处理器<br/>
	 * 
	 * @param is
	 *            输入流
	 * @param version
	 *            版本
	 * @param readersByName
	 *            按照表名匹配的Sheet读取处理器集合[处理器实例不可复用]
	 */
	public static void readBySheetNames(InputStream is, ExcelVersion version, Map<String, SheetReader> readersByName) {
		getReader(version).readBySheetNames(is, readersByName);
	}

	/**
	 * 根据版本获取Excel读取器
	 * 
	 * @param version
	 *            Excel版本
	 * @return
	 */
	private static CoreReader getReader(ExcelVersion version) {
		if (version == null) {
			return new XlsxReader();
		} else if (ExcelVersion.XLS.equals(version)) {
			return new XlsReader();
		} else {
			return new XlsxReader();
		}
	}

}

/**
 * 路径转换
 * 
 * @author Frank
 *
 */
class PathParser {

	/**
	 * 版本
	 */
	private ExcelVersion version;

	/**
	 * 文件输入流
	 */
	private FileInputStream fis;

	/**
	 * 
	 * @param excelPath
	 *            Excel路径
	 */
	public PathParser(String excelPath) {
		if (excelPath == null) {
			throw new IllegalArgumentException("excelPath is null");
		}
		if (!new File(excelPath).exists()) {
			throw new IllegalArgumentException("excelPath[" + excelPath + "] is not exists");
		}
		if (excelPath.endsWith(ExcelVersion.XLSX.getSuffix())) {
			version = ExcelVersion.XLSX;
		} else if (excelPath.endsWith(ExcelVersion.XLS.getSuffix())) {
			version = ExcelVersion.XLS;
		} else {
			throw new IllegalArgumentException("excelPath[" + excelPath + "] is not excel");
		}
		try {
			fis = new FileInputStream(excelPath);
		} catch (IOException e) {
			throw new IllegalArgumentException(e);
		}
	}

	public ExcelVersion getVersion() {
		return version;
	}

	public InputStream getInputStream() {
		return fis;
	}
}

/**
 * 结果全部保存到内存中统一返回的数据处理器
 * 
 * @author Frank
 *
 * @param <D>
 */
class ReturnSchemaDataReader<D> implements SheetDataHandler<D> {

	/**
	 * 所有数据
	 */
	private List<D> list = new ArrayList<>();
	/**
	 * 数据Class
	 */
	private Class<D> dataClass;

	/**
	 * 
	 * @param dataClass
	 *            数据Class
	 */
	public ReturnSchemaDataReader(Class<D> dataClass) {
		this.dataClass = dataClass;
	}

	@Override
	public void handle(int rowIndex, D data) {
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
