package cn.emay.excel.read;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import cn.emay.excel.common.ExcelVersion;
import cn.emay.excel.common.schema.base.SheetSchema;
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
	 * 从文件中读取Excel表格第一个单元格<br/>
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
	 * 从文件中按照表格名字读取Excel表格<br/>
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
	 * 从文件中按照表格序号读取Excel表格<br/>
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
	 * 从文件中读取Excel表格第一个单元格<br/>
	 * dataClass 实现了@ExcelSheet注解,其字段实现了@ExcelColumn注解
	 * 
	 * @param excelPath
	 *            路径
	 * @param dataReader
	 *            数据处理器
	 * @return
	 */
	public static <D> void readFirstSheet(String excelPath, SheetDataHandler<D> dataReader) {
		readBySheetIndex(excelPath, 0, dataReader);
	}

	/**
	 * 从文件中按照表格名字读取Excel表格<br/>
	 * dataClass实现了@ExcelSheet注解,其字段实现了@ExcelColumn注解
	 * 
	 * @param excelPath
	 *            路径
	 * @param sheetName
	 *            Sheet页名字
	 * @param dataReader
	 *            数据处理器
	 */
	public static <D> void readBySheetName(String excelPath, String sheetName, SheetDataHandler<D> dataReader) {
		SchemaSheetReader<D> handler = new SchemaSheetReader<D>(new SheetSchema<D>(dataReader.getDataClass()), dataReader);
		readBySheetName(excelPath, sheetName, handler);
	}

	/**
	 * 从文件中按照表格序号读取Excel表格<br/>
	 * dataClass实现了@ExcelSheet注解,其字段实现了@ExcelColumn注解
	 * 
	 * @param excelPath
	 *            路径
	 * @param sheetIndex
	 *            Sheet Index
	 * @param dataReader
	 *            数据处理器
	 */
	public static <D> void readBySheetIndex(String excelPath, int sheetIndex, SheetDataHandler<D> dataReader) {
		SchemaSheetReader<D> handler = new SchemaSheetReader<D>(new SheetSchema<D>(dataReader.getDataClass()), dataReader);
		readBySheetIndex(excelPath, sheetIndex, handler);
	}

	/**
	 * 从文件中读取Excel表格第一个单元格<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param customSchemaReader
	 *            自定义的表格定义读取器
	 * @return
	 */
	public static <D> void readFirstSheet(String excelPath,  SchemaSheetDataHandler<D> schemaSheetDataHandler) {
		readBySheetIndex(excelPath, 0, schemaSheetDataHandler);
	}

	/**
	 * 从文件中按照表格名字读取Excel表格<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param sheetName
	 *            Sheet页名字
	 * @param customSchemaReader
	 *            自定义的表格定义读取器
	 */
	public static <D> void readBySheetName(String excelPath, String sheetName,  SchemaSheetDataHandler<D> schemaSheetDataHandler) {
		SchemaSheetReader<D> handler = new SchemaSheetReader<D>(schemaSheetDataHandler.getSheetSchema(), schemaSheetDataHandler);
		readBySheetName(excelPath, sheetName, handler);
	}

	/**
	 * 从文件中按照表格序号读取Excel表格<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param sheetIndex
	 *            Sheet Index
	 * @param customSchemaReader
	 *            自定义的表格定义读取器
	 */
	public static <D> void readBySheetIndex(String excelPath, int sheetIndex, SchemaSheetDataHandler<D> schemaSheetDataHandler) {
		SchemaSheetReader<D> handler = new SchemaSheetReader<D>(schemaSheetDataHandler.getSheetSchema(), schemaSheetDataHandler);
		readBySheetIndex(excelPath, sheetIndex, handler);
	}

	/**
	 * 从文件中读取Excel表格第一个sheet页<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param handler
	 *            Sheet读取处理器[处理器实例不可复用]
	 */
	public static void readFirstSheet(String excelPath, SheetReader handler) {
		readBySheetIndex(excelPath, 0, handler);
	}

	/**
	 * 从文件中读取Excel表格<br/>
	 * 按照表格序号读取<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param sheetIndex
	 *            Sheet Index
	 * @param handler
	 *            Sheet读取处理器[处理器实例不可复用]
	 */
	public static void readBySheetIndex(String excelPath, int sheetIndex, SheetReader handler) {
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
	public static void readBySheetIndexs(String excelPath, Map<Integer, SheetReader> handlersByIndex) {
		readPath(excelPath, new FileHandler() {
			@Override
			public void readHandler(InputStream is, ExcelVersion version) {
				readBySheetIndexs(is, version, handlersByIndex);
			}
		});
	}

	/**
	 * 从文件中读取Excel表格<br/>
	 * 按照表格名字读取<br/>
	 * 
	 * @param excelPath
	 *            路径
	 * @param sheetName
	 *            Sheet页名字
	 * @param handler
	 *            Sheet读取处理器[处理器实例不可复用]
	 */
	public static void readBySheetName(String excelPath, String sheetName, SheetReader handler) {
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
	public static void readBySheetNames(String excelPath, Map<String, SheetReader> handlersByName) {
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
	public static void readFirstSheet(InputStream is, ExcelVersion version, SheetReader handler) {
		readBySheetIndex(is, version, 0, handler);
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
	 * @param handler
	 *            Sheet读取处理器[处理器实例不可复用]
	 */
	public static void readBySheetIndex(InputStream is, ExcelVersion version, int sheetIndex, SheetReader handler) {
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
	public static void readBySheetIndexs(InputStream is, ExcelVersion version, Map<Integer, SheetReader> handlersByIndex) {
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
	 * 按照表格名字读取<br/>
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
	public static void readBySheetName(InputStream is, ExcelVersion version, String sheetName, SheetReader handler) {
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
	public static void readBySheetNames(InputStream is, ExcelVersion version, Map<String, SheetReader> handlersByName) {
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
	
	private Class<D> dataClass;
	
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
