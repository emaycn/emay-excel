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
import cn.emay.excel.read.core.XlsReader;
import cn.emay.excel.read.core.XlsxReader;
import cn.emay.excel.read.handler.SheetReadHandler;
import cn.emay.excel.read.handler.relize.DataReader;
import cn.emay.excel.read.handler.relize.SheetReadBySchemaHandler;

/**
 * Excel读<br/>
 * XLSX统一采用SAX方式读取
 * 
 * @author Frank
 *
 */
public class ExcelReader {

	/*--------------------SCHEMA----------------------------*/

	/**
	 * 从文件中读取Excel表格第一个单元格<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param dataReader
	 *            Sheet读取处理器[处理器实例不可复用]
	 * @return
	 */
	public static <D> List<D> readFirstSheetWithSchema(String excelPath, Class<D> schemaClass) {
		return readBySheetIndexWithSchema(excelPath, 0, schemaClass);
	}

	/**
	 * 从文件中读取Excel表格<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param sheetName
	 *            Sheet页名字
	 * @param dataReader
	 *            Sheet读取处理器[处理器实例不可复用]
	 * @return
	 */
	public static <D> List<D> readBySheetNameWithSchema(String excelPath, String sheetName, Class<D> schemaClass) {
		if (sheetName == null) {
			throw new IllegalArgumentException("sheetName is null");
		}
		if (schemaClass == null) {
			throw new IllegalArgumentException("schemaClass is null");
		}
		NornamDataReader<D> nor = new NornamDataReader<D>(schemaClass);
		readBySheetNameWithReader(excelPath, sheetName, nor);
		return nor.getResult();
	}

	/**
	 * 从文件中读取Excel表格<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param sheetIndex
	 *            Sheet Index
	 * @param dataReader
	 *            Sheet读取处理器[处理器实例不可复用]
	 * @return
	 */
	public static <D> List<D> readBySheetIndexWithSchema(String excelPath, int sheetIndex, Class<D> schemaClass) {
		if (schemaClass == null) {
			throw new IllegalArgumentException("schemaClass is null");
		}
		NornamDataReader<D> nor = new NornamDataReader<D>(schemaClass);
		readBySheetIndexWithReader(excelPath, sheetIndex, nor);
		return nor.getResult();
	}

	/*----------DataReader---------------*/

	/**
	 * 从文件中读取Excel表格第一个sheet页<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param dataReader
	 *            Sheet读取处理器[处理器实例不可复用]
	 */
	public static void readFirstSheetWithReader(String excelPath, DataReader<?> dataReader) {
		readBySheetIndexWithReader(excelPath, 0, dataReader);
	}

	/**
	 * 从文件中读取Excel表格<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param sheetIndex
	 *            Sheet Index
	 * @param dataReader
	 *            Sheet读取处理器[处理器实例不可复用]
	 */
	public static void readBySheetIndexWithReader(String excelPath, int sheetIndex, DataReader<?> dataReader) {
		if (dataReader == null) {
			throw new IllegalArgumentException("dataReader is null");
		}
		Map<Integer, DataReader<?>> handlers = new HashMap<>(1);
		handlers.put(sheetIndex, dataReader);
		readBySheetIndexsWithReader(excelPath, handlers);
	}

	/**
	 * 从文件中读取Excel表格<br/>
	 * 按照表序号匹配读取处理器<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param dataReadersByIndex
	 *            按照Index匹配的Sheet读取处理器集合[处理器实例不可复用]
	 */
	public static void readBySheetIndexsWithReader(String excelPath, Map<Integer, DataReader<?>> dataReadersByIndex) {
		if (dataReadersByIndex == null) {
			throw new IllegalArgumentException("dataReadersByIndex is null");
		}
		Map<Integer, SheetReadHandler> handlers = new HashMap<>(dataReadersByIndex.size());
		for (Integer sheetIndex : dataReadersByIndex.keySet()) {
			DataReader<?> handler = dataReadersByIndex.get(sheetIndex);
			if (handler == null) {
				continue;
			}
			@SuppressWarnings({ "unchecked", "rawtypes" })
			SheetReadBySchemaHandler<?> schandler = new SheetReadBySchemaHandler(handler);
			handlers.put(sheetIndex, schandler);
		}
		readBySheetIndexs(excelPath, handlers);
	}

	/**
	 * 从文件中读取Excel表格<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param sheetName
	 *            Sheet页名字
	 * @param dataReader
	 *            Sheet读取处理器[处理器实例不可复用]
	 */
	public static void readBySheetNameWithReader(String excelPath, String sheetName, DataReader<?> dataReader) {
		if (sheetName == null) {
			throw new IllegalArgumentException("sheetName is null");
		}
		if (dataReader == null) {
			throw new IllegalArgumentException("dataReader is null");
		}
		Map<String, DataReader<?>> handlersByName = new HashMap<>(1);
		handlersByName.put(sheetName, dataReader);
		readBySheetNamesWithReader(excelPath, handlersByName);
	}

	/**
	 * 从文件中读取Excel表格<br/>
	 * 按照表名匹配读取处理器<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param dataReadersByName
	 *            按照表名匹配的Sheet读取处理器集合[处理器实例不可复用]
	 */
	public static void readBySheetNamesWithReader(String excelPath, Map<String, DataReader<?>> dataReadersByName) {
		if (dataReadersByName == null) {
			throw new IllegalArgumentException("dataReadersByName is null");
		}
		Map<String, SheetReadHandler> handlersByName = new HashMap<>(dataReadersByName.size());
		for (String sheetName : dataReadersByName.keySet()) {
			DataReader<?> handler = dataReadersByName.get(sheetName);
			if (handler == null) {
				continue;
			}
			@SuppressWarnings({ "unchecked", "rawtypes" })
			SheetReadBySchemaHandler<?> schandler = new SheetReadBySchemaHandler(handler);
			handlersByName.put(sheetName, schandler);
		}
		readBySheetNames(excelPath, handlersByName);
	}

	/*---------------------BASE-----------------------------*/

	/**
	 * 从文件中读取Excel表格第一个sheet页<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param handler
	 *            Sheet读取处理器[处理器实例不可复用]
	 */
	public static void readFirstSheet(String excelPath, SheetReadHandler handler) {
		readBySheetIndex(excelPath, 0, handler);
	}

	/**
	 * 从文件中读取Excel表格<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param sheetIndex
	 *            Sheet Index
	 * @param handler
	 *            Sheet读取处理器[处理器实例不可复用]
	 */
	public static void readBySheetIndex(String excelPath, int sheetIndex, SheetReadHandler handler) {
		readPath(excelPath, new FileHandler() {
			@Override
			public void readHandler(InputStream is, ExcelVersion version) {
				readBySheetIndex(is, version, sheetIndex, handler);
			}
		});
	}

	/**
	 * 从文件中读取Excel表格<br/>
	 * 按照表序号匹配读取处理器<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param handlersByIndex
	 *            按照Index匹配的Sheet读取处理器集合[处理器实例不可复用]
	 */
	public static void readBySheetIndexs(String excelPath, Map<Integer, SheetReadHandler> handlersByIndex) {
		readPath(excelPath, new FileHandler() {
			@Override
			public void readHandler(InputStream is, ExcelVersion version) {
				readBySheetIndexs(is, version, handlersByIndex);
			}
		});
	}

	/**
	 * 从文件中读取Excel表格<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param sheetName
	 *            Sheet页名字
	 * @param handler
	 *            Sheet读取处理器[处理器实例不可复用]
	 */
	public static void readBySheetName(String excelPath, String sheetName, SheetReadHandler handler) {
		readPath(excelPath, new FileHandler() {
			@Override
			public void readHandler(InputStream is, ExcelVersion version) {
				readBySheetName(is, version, sheetName, handler);
			}
		});
	}

	/**
	 * 从文件中读取Excel表格<br/>
	 * 按照表名匹配读取处理器<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param handlersByName
	 *            按照表名匹配的Sheet读取处理器集合[处理器实例不可复用]
	 */
	public static void readBySheetNames(String excelPath, Map<String, SheetReadHandler> handlersByName) {
		readPath(excelPath, new FileHandler() {
			@Override
			public void readHandler(InputStream is, ExcelVersion version) {
				readBySheetNames(is, version, handlersByName);
			}
		});
	}

	/**
	 * 统一读取方法
	 * 
	 * @param excelPath
	 *            路径
	 * @param read
	 *            读取方法
	 */
	private static void readPath(String excelPath, FileHandler read) {
		if (excelPath == null) {
			throw new IllegalArgumentException("excelPath is null");
		}
		if (!new File(excelPath).exists()) {
			throw new IllegalArgumentException("excelPath[" + excelPath + "] is not exists");
		}
		ExcelVersion version;
		if (excelPath.endsWith(ExcelVersion.XLSX.getSuffix())) {
			version = ExcelVersion.XLSX;
		} else if (excelPath.endsWith(ExcelVersion.XLS.getSuffix())) {
			version = ExcelVersion.XLS;
		} else {
			throw new IllegalArgumentException("excelPath[" + excelPath + "] is not excel");
		}
		FileInputStream fis = null;
		try {
			fis = new FileInputStream(excelPath);
			read.readHandler(fis, version);
		} catch (IOException e) {
			throw new IllegalArgumentException(e);
		} finally {
			if (fis != null) {
				try {
					fis.close();
				} catch (IOException e) {
					throw new IllegalArgumentException(e);
				}
			}
		}
	}

	/**
	 * 从文件中读取Excel表格第一个sheet页<br/>
	 * 
	 * @param is
	 *            输入流
	 * @param version
	 *            版本
	 * @param handler
	 *            Sheet读取处理器[处理器实例不可复用]
	 */
	public static void readFirstSheet(InputStream is, ExcelVersion version, SheetReadHandler handler) {
		readBySheetIndex(is, version, 0, handler);
	}

	/**
	 * 从输入流中读取Excel表格<br/>
	 * 
	 * @param is
	 *            输入流
	 * @param version
	 *            版本
	 * @param sheetIndex
	 *            Sheet Index
	 * @param handler
	 *            Sheet读取处理器[处理器实例不可复用]
	 */
	public static void readBySheetIndex(InputStream is, ExcelVersion version, int sheetIndex, SheetReadHandler handler) {
		readInputStream(is, version, new InputStreamHandler() {
			@Override
			public void readHandler() {
				XlsxReader.readBySheetIndex(is, sheetIndex, handler);
			}
		}, new InputStreamHandler() {
			@Override
			public void readHandler() {
				XlsReader.readBySheetIndex(is, sheetIndex, handler);
			}
		});
	}

	/**
	 * 从输入流中读取Excel表格<br/>
	 * 按照表序号匹配读取处理器<br/>
	 * 
	 * @param is
	 *            输入流
	 * @param version
	 *            版本
	 * @param handlersByIndex
	 *            按照Index匹配的Sheet读取处理器集合[处理器实例不可复用]
	 */
	public static void readBySheetIndexs(InputStream is, ExcelVersion version, Map<Integer, SheetReadHandler> handlersByIndex) {
		readInputStream(is, version, new InputStreamHandler() {
			@Override
			public void readHandler() {
				XlsxReader.readBySheetIndexs(is, handlersByIndex);
			}
		}, new InputStreamHandler() {
			@Override
			public void readHandler() {
				XlsReader.readBySheetIndexs(is, handlersByIndex);
			}
		});
	}

	/**
	 * 从输入流中读取Excel表格<br/>
	 * 
	 * @param is
	 *            输入流
	 * @param version
	 *            版本
	 * @param sheetName
	 *            Sheet页名字
	 * @param handler
	 *            Sheet读取处理器[处理器实例不可复用]
	 */
	public static void readBySheetName(InputStream is, ExcelVersion version, String sheetName, SheetReadHandler handler) {
		readInputStream(is, version, new InputStreamHandler() {
			@Override
			public void readHandler() {
				XlsxReader.readBySheetName(is, sheetName, handler);
			}
		}, new InputStreamHandler() {
			@Override
			public void readHandler() {
				XlsReader.readBySheetName(is, sheetName, handler);
			}
		});
	}

	/**
	 * 从输入流中读取Excel表格<br/>
	 * 按照表名匹配读取处理器<br/>
	 * 
	 * @param is
	 *            输入流
	 * @param version
	 *            版本
	 * @param handlersByName
	 *            按照表名匹配的Sheet读取处理器集合[处理器实例不可复用]
	 */
	public static void readBySheetNames(InputStream is, ExcelVersion version, Map<String, SheetReadHandler> handlersByName) {
		readInputStream(is, version, new InputStreamHandler() {
			@Override
			public void readHandler() {
				XlsxReader.readBySheetNames(is, handlersByName);
			}
		}, new InputStreamHandler() {
			@Override
			public void readHandler() {
				XlsReader.readBySheetNames(is, handlersByName);
			}
		});
	}

	/**
	 * 统一读取方法
	 * 
	 * @param version
	 *            版本
	 * @param xlsx
	 *            Xlsx读取方法
	 * @param xls
	 *            Xls读取方法
	 */
	private static void readInputStream(InputStream is, ExcelVersion version, InputStreamHandler xlsx, InputStreamHandler xls) {
		if (version == null) {
			throw new IllegalArgumentException("ExcelVersion is null");
		}
		switch (version) {
		case XLSX:
			xlsx.readHandler();
			break;
		case XLS:
			xls.readHandler();
			break;
		default:
			break;
		}
	}

}

/**
 * 读取统一处理方法
 * 
 * @author Frank
 *
 */
interface InputStreamHandler {

	/**
	 * 读取操作
	 */
	void readHandler();
}

/**
 * 读取统一处理方法
 * 
 * @author Frank
 *
 */
interface FileHandler {

	/**
	 * 读取操作
	 * 
	 * @param is
	 * @param version
	 */
	void readHandler(InputStream is, ExcelVersion version);
}

/**
 * 内置的数据处理器
 * 
 * @author Frank
 *
 * @param <D>
 */
class NornamDataReader<D> implements DataReader<D> {

	private List<D> list = new ArrayList<>();

	private Class<D> clazz;

	public NornamDataReader(Class<D> clazz) {
		this.clazz = clazz;
	}

	@Override
	public void handlerRowData(int rowIndex, D data) {
		if (data != null) {
			list.add(data);
		}
	}

	@Override
	public Class<D> getShemaClass() {
		return clazz;
	}

	public List<D> getResult() {
		return list;
	}

}
