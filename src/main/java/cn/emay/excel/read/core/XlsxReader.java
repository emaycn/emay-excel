package cn.emay.excel.read.core;

import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import cn.emay.excel.read.handler.SheetReadHandler;

/**
 * XLSX读处理器
 * 
 * @author Frank
 *
 */
public class XlsxReader {

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
	public static void readBySheetIndex(InputStream is, int sheetIndex, SheetReadHandler handler) {
		if (handler == null) {
			throw new IllegalArgumentException("handler is null");
		}
		if (sheetIndex < 0) {
			throw new IllegalArgumentException("sheetIndex begin with 0 , and must bigger than 0");
		}
		Map<Integer, SheetReadHandler> handlersByName = new HashMap<>(1);
		handlersByName.put(sheetIndex, handler);
		readBySheetIndexs(is, handlersByName);
	}

	/**
	 * 从输入流中读取Excel表格<br/>
	 * 
	 * @param is
	 *            输入流
	 * @param handlers
	 *            按照Index匹配的Sheet读取处理器集合[数组index=sheet index]
	 */
	public static void readByOrder(InputStream is, SheetReadHandler... handlers) {
		if (handlers == null || handlers.length == 0) {
			throw new IllegalArgumentException("handlers is null");
		}
		Map<Integer, SheetReadHandler> handlersByIndex = new HashMap<>(handlers.length);
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
	public static void readBySheetIndexs(InputStream is, Map<Integer, SheetReadHandler> handlersByIndex) {
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
	public static void readBySheetName(InputStream is, String sheetName, SheetReadHandler handler) {
		if (handler == null) {
			throw new IllegalArgumentException("handler is null");
		}
		if (sheetName == null) {
			throw new IllegalArgumentException("sheetName is null");
		}
		Map<String, SheetReadHandler> handlersByName = new HashMap<>(1);
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
	public static void readBySheetNames(InputStream is, Map<String, SheetReadHandler> handlersByName) {
		if (handlersByName == null || handlersByName.size() == 0) {
			throw new IllegalArgumentException("handlers is null");
		}
		read(is, null, handlersByName);
	}

	/**
	 * 从输入流中读取Excel表格<br/>
	 * SAX方式<br/>
	 * 所有处理器依次进行匹配
	 * 
	 * @param is
	 *            输入流
	 * @param handlersByIndex
	 *            按照Index匹配的Sheet读取处理器集合
	 * @param handlersByName
	 *            按照表名匹配的Sheet读取处理器集合
	 */
	public static void read(InputStream is, Map<Integer, SheetReadHandler> handlersByIndex, Map<String, SheetReadHandler> handlersByName) {
		if (is == null) {
			throw new IllegalArgumentException("InputStream is null");
		}
		OPCPackage opcPackage = null;
		try {
			opcPackage = OPCPackage.open(is);
			XSSFReader xssfReader = new XSSFReader(opcPackage);
			StylesTable stylesTable = xssfReader.getStylesTable();
			SharedStringsTable sst = xssfReader.getSharedStringsTable();
			XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
			int sheetIndex = 0;
			while (iter.hasNext()) {
				InputStream sheet = null;
				try {
					sheet = iter.next();
					String sheetName = iter.getSheetName();
					SheetReadHandler readHander = null;
					if(handlersByIndex != null) {
						readHander = handlersByIndex.get(sheetIndex);
					}
					if(readHander == null && handlersByName != null) {
						readHander = handlersByName.get(sheetName);
					}
					if (readHander == null) {
						continue;
					}
					InputSource sheetSource = new InputSource(sheet);
					SAXParserFactory saxFactory = SAXParserFactory.newInstance();
					SAXParser saxParser = saxFactory.newSAXParser();
					XMLReader sheetParser = saxParser.getXMLReader();
					XlsxSheetHandler mxHandler = new XlsxSheetHandler(stylesTable, sst,sheetIndex ,sheetName,readHander);
					sheetParser.setContentHandler(mxHandler);
					sheetParser.parse(sheetSource);
				} catch (XlsxStopReadException e) {
					// 本sheet停止读取
					continue;
				} catch (IOException e) {
					throw new IllegalArgumentException(e);
				} catch (SAXException e) {
					throw new IllegalArgumentException(e);
				} catch (ParserConfigurationException e) {
					throw new IllegalArgumentException(e);
				} finally {
					sheetIndex++;
					if (sheet != null) {
						sheet.close();
					}
				}
			}
		} catch (

		InvalidFormatException e) {
			throw new IllegalArgumentException(e);
		} catch (IOException e) {
			throw new IllegalArgumentException(e);
		} catch (OpenXML4JException e) {
			throw new IllegalArgumentException(e);
		} finally {
			if (opcPackage != null) {
				try {
					opcPackage.close();
				} catch (IOException e) {
					throw new IllegalArgumentException(e);
				}
			}
			try {
				is.close();
			} catch (IOException e) {
				throw new IllegalArgumentException(e);
			}
		}
	}

}
