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
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

import cn.emay.excel.read.reader.SheetReader;

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
	public static void readBySheetIndex(InputStream is, int sheetIndex, SheetReader handler) {
		if (handler == null) {
			throw new IllegalArgumentException("handler is null");
		}
		if (sheetIndex < 0) {
			throw new IllegalArgumentException("sheetIndex begin with 0 , and must bigger than 0");
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
	 *            按照Index匹配的Sheet读取处理器集合[数组index=sheet index]
	 */
	public static void readByOrder(InputStream is, SheetReader... handlers) {
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
	public static void readBySheetIndexs(InputStream is, Map<Integer, SheetReader> handlersByIndex) {
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
	public static void readBySheetName(InputStream is, String sheetName, SheetReader handler) {
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
	public static void readBySheetNames(InputStream is, Map<String, SheetReader> handlersByName) {
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
	public static void read(InputStream is, Map<Integer, SheetReader> handlersByIndex, Map<String, SheetReader> handlersByName) {
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
					SheetReader readHander = null;
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

/**
 * SAX方式读取ExcelSheet页
 * 
 * @author Frank
 *
 */
class XlsxSheetHandler extends DefaultHandler {

	/**
	 * Sheet读处理器
	 */
	private SheetReader handler;
	/**
	 * SAX值列表
	 */
	private SharedStringsTable sst;
	/**
	 * Excel样式集
	 */
	private StylesTable stylesTable;
	/**
	 * 当前处理的sheet的Index
	 */
	private int sheetIndex;
	/**
	 * 当前处理的sheet的名字
	 */
	private String sheetName;
	/**
	 * 当前单元格值
	 */
	private String lastContents;
	/**
	 * 当前单元格是否是字符串
	 */
	private boolean nextIsS;
	/**
	 * 当前单元格列序号
	 */
	private int currColumnIndex = 0;
	/**
	 * 当前行号
	 */
	private int currRowIndex = 0;
	/**
	 * 上一行行号
	 */
	private int preRowIndex = -1;
	/**
	 * 单元格数据类型
	 */
	private int formatIndex = -1;
	/**
	 * 开始读取的行号
	 */
	private int startReadRowIndex = 0;
	/**
	 * 结束读取的行号
	 */
	private int endReadRowIndex = 0;

	public XlsxSheetHandler(StylesTable stylesTable, SharedStringsTable sst, int sheetIndex, String sheetName, SheetReader handler) {
		this.sst = sst;
		this.handler = handler;
		this.stylesTable = stylesTable;
		this.startReadRowIndex = handler.getStartReadRowIndex();
		this.endReadRowIndex = handler.getEndReadRowIndex();
		this.sheetIndex = sheetIndex;
		this.sheetName = sheetName;
	}

	@Override
	public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
		if ("row".equals(name)) {
			currRowIndex = Integer.valueOf(attributes.getValue("r")) - 1;
		}
		if (endReadRowIndex >= 0 && currRowIndex > endReadRowIndex) {
			handler.endRow(preRowIndex);
			handler.end(sheetIndex, sheetName);
			// 停止读取
			throw new XlsxStopReadException();
		}
		if (startReadRowIndex > currRowIndex) {
			return;
		}
		if ("row".equals(name)) {
			if (preRowIndex != -1) {
				handler.endRow(preRowIndex);
			}
			handler.beginRow(currRowIndex);
			preRowIndex = currRowIndex;
		}
		if ("c".equals(name)) {
			String coordinate = attributes.getValue("r");
			String columnIndexStr = coordinate.replaceAll("\\d+", "");
			currColumnIndex = (int) XlsxWord26.decode(columnIndexStr);
			String cellType = attributes.getValue("t");
			if (cellType != null && cellType.equals("s")) {
				nextIsS = true;
				formatIndex = -1;
			} else {
				nextIsS = false;
				String cellStyleStr = attributes.getValue("s");
				if (cellStyleStr != null) {
					int styleIndex = Integer.parseInt(cellStyleStr);
					XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
					formatIndex = style.getDataFormat();
				} else {
					formatIndex = -1;
				}
			}
		}
		lastContents = "";
	}

	@Override
	public void endElement(String uri, String localName, String name) throws SAXException {
		if (startReadRowIndex > currRowIndex) {
			return;
		}
		if (nextIsS) {
			int idx = Integer.parseInt(lastContents);
			lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
			nextIsS = false;
		}
		if ("v".equals(name)) {
			handler.handleXlsxCell(currRowIndex, currColumnIndex, formatIndex, lastContents);
		}
		// 兼容SXSSFWorkbook写的文件
		if ("t".equals(name)) {
			handler.handleXlsxCell(currRowIndex, currColumnIndex, formatIndex, lastContents);
		}
	}

	@Override
	public void startDocument() throws SAXException {
		handler.begin(sheetIndex, sheetName);
	}

	@Override
	public void endDocument() throws SAXException {
		handler.endRow(currRowIndex);
		handler.end(sheetIndex, sheetName);
	}

	@Override
	public void characters(char[] ch, int start, int length) {
		lastContents += new String(ch, start, length);
	}

}

/**
 * 大写字母26进制
 * 
 * @author Frank
 *
 */
class XlsxWord26 {

	/**
	 * 进制数
	 */
	private static long HEX;

	/**
	 * 进制字符集
	 */
	private static char[] CHARS;

	/**
	 * 进制对应的数字集
	 */
	private static Integer[] INTS;

	/**
	 * 初始化
	 * 
	 */
	static {
		CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".toCharArray();
		HEX = CHARS.length;
		INTS = new Integer[CHARS.length];
		for (int i = 0; i < HEX; i++) {
			INTS[i] = i;
		}
	}

	/**
	 * 根据字符拿到数字
	 * 
	 * @param cha
	 * @return
	 */
	private static Integer getIndexChar(char cha) {
		int index = -1;
		for (int i = 0; i < HEX; i++) {
			if (CHARS[i] == cha) {
				index = i;
				break;
			}
		}
		if (index == -1) {
			return null;
		}
		return INTS[index];
	}

	/**
	 * 解码
	 * 
	 * @param inhex
	 * @return
	 */
	public static long decode(String inhex) {
		if (inhex == null || inhex.trim().length() == 0) {
			return 0L;
		}
		long result = 0L;
		char[] chars = inhex.toCharArray();
		for (int i = 0; i < chars.length; i++) {
			char cha = chars[i];
			Integer ind = getIndexChar(cha);
			if (ind == null) {
				return 0L;
			}
			result += Math.pow(HEX, chars.length - i - 1) * ind;
		}
		return result;
	}

}

/**
 * 
 * SAX读取方式停止异常
 * 
 * @author Frank
 *
 */
class XlsxStopReadException extends RuntimeException {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	public XlsxStopReadException() {
		super();
	}

}