package cn.emay.excel.read.core;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.Map;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
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
 * 
 * @author Frank
 *
 */
public class XlsxReader extends BaseReader {

	@Override
	public void read(InputStream is, Map<Integer, SheetReader> handlersByIndex, Map<String, SheetReader> handlersByName) {
		if (is == null) {
			throw new IllegalArgumentException("InputStream is null");
		}
		try {
			OPCPackage opcPackage = OPCPackage.open(is);
			readByOPCPackage(opcPackage, handlersByIndex, handlersByName);
		} catch (Exception e) {
			throw new IllegalArgumentException(e);
		} finally {
			try {
				is.close();
			} catch (IOException e) {
				throw new IllegalArgumentException(e);
			}
		}

	}

	@Override
	public void read(File file, Map<Integer, SheetReader> handlersByIndex, Map<String, SheetReader> handlersByName) {
		if (file == null) {
			throw new IllegalArgumentException("file is null");
		}
		try {
			OPCPackage opcPackage = OPCPackage.open(file);
			readByOPCPackage(opcPackage, handlersByIndex, handlersByName);
		} catch (Exception e) {
			throw new IllegalArgumentException(e);
		}
	}

	private void readByOPCPackage(OPCPackage opcPackage, Map<Integer, SheetReader> handlersByIndex, Map<String, SheetReader> handlersByName) {
		if (opcPackage == null) {
			throw new IllegalArgumentException("opcPackage is null");
		}
		try {
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
					if (handlersByIndex != null) {
						readHander = handlersByIndex.get(sheetIndex);
					}
					if (readHander == null && handlersByName != null) {
						readHander = handlersByName.get(sheetName);
					}
					if (readHander == null) {
						continue;
					}
					InputSource sheetSource = new InputSource(sheet);
					SAXParserFactory saxFactory = SAXParserFactory.newInstance();
					SAXParser saxParser = saxFactory.newSAXParser();
					XMLReader sheetParser = saxParser.getXMLReader();
					XlsxSheetHandler mxHandler = new XlsxSheetHandler(stylesTable, sst, sheetIndex, sheetName, readHander);
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
		} catch (InvalidFormatException e) {
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
		}
	}

	/**
	 * 
	 * SAX读取方式停止异常
	 * 
	 * @author Frank
	 *
	 */
	public static class XlsxStopReadException extends RuntimeException {

		/**
		 * 
		 */
		private static final long serialVersionUID = 1L;

		public XlsxStopReadException() {
			super();
		}

	}

	/**
	 * 数据类型
	 * 
	 * @author Frank
	 *
	 */
	public static enum DataType {
		/**
		 * 布尔
		 */
		BOOL,
		/**
		 * 错误
		 */
		ERROR,
		/**
		 * 公式
		 */
		FORMULA,
		/**
		 * String
		 */
		INLINESTR,
		/**
		 * String
		 */
		SSTINDEX,
		/**
		 * 数字，包括日期
		 */
		NUMBER,
	}

	/**
	 * 处理器
	 * 
	 * @author Frank
	 *
	 */
	public static class XlsxSheetHandler extends DefaultHandler {

		private StylesTable stylesTable;
		private SharedStringsTable sharedStringsTable;
		private final DataFormatter formatter;

		private String sheetName;
		private int sheetIndex;

		private SheetReader handler;
		private int startReadRowIndex = 0;
		private int endReadRowIndex = 0;

		private short formatIndex;
		private String formatString;
		private DataType nextDataType;
		private StringBuffer value;

		private int currColumnIndex = 0;
		private int currRowIndex = 0;
		private int preRowIndex = -1;

		private boolean vIsOpen;

		public XlsxSheetHandler(StylesTable styles, SharedStringsTable strings, int sheetIndex, String sheetName, SheetReader handler) {
			this.stylesTable = styles;
			this.sharedStringsTable = strings;
			this.formatter = new DataFormatter();

			this.value = new StringBuffer();
			this.nextDataType = DataType.NUMBER;

			this.handler = handler;
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

			if ("inlineStr".equals(name) || "v".equals(name) || "t".equals(name)) {
				vIsOpen = true;
				value.setLength(0);
			} else if ("c".equals(name)) {
				String coordinate = attributes.getValue("r");
				String columnIndexStr = coordinate.replaceAll("\\d+", "");
				currColumnIndex = (int) decode(columnIndexStr);

				this.nextDataType = DataType.NUMBER;
				this.formatIndex = -1;
				this.formatString = null;
				String cellType = attributes.getValue("t");
				String cellStyleStr = attributes.getValue("s");
				if ("b".equals(cellType)) {
					nextDataType = DataType.BOOL;
				} else if ("e".equals(cellType)) {
					nextDataType = DataType.ERROR;
				} else if ("inlineStr".equals(cellType)) {
					nextDataType = DataType.INLINESTR;
				} else if ("s".equals(cellType)) {
					nextDataType = DataType.SSTINDEX;
				} else if ("str".equals(cellType)) {
					nextDataType = DataType.FORMULA;
				} else if (cellStyleStr != null) {
					int styleIndex = Integer.parseInt(cellStyleStr);
					XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
					this.formatIndex = style.getDataFormat();
					this.formatString = style.getDataFormatString();
					if (this.formatString == null) {
						this.formatString = BuiltinFormats.getBuiltinFormat(this.formatIndex);
					}
				}
			}

		}

		@Override
		public void endElement(String uri, String localName, String name) throws SAXException {
			if (startReadRowIndex > currRowIndex) {
				return;
			}
			String thisStr = null;
			if ("v".equals(name) || "t".equals(name)) {
				switch (nextDataType) {
				case BOOL:
					char first = value.charAt(0);
					thisStr = first == '0' ? "FALSE" : "TRUE";
					break;
				case FORMULA:
					thisStr = value.toString();
					break;
				case INLINESTR:
					XSSFRichTextString rtsi = new XSSFRichTextString(value.toString());
					thisStr = rtsi.toString();
					break;
				case SSTINDEX:
					String sstIndex = value.toString();
					try {
						int idx = Integer.parseInt(sstIndex);
						XSSFRichTextString rtss = new XSSFRichTextString(sharedStringsTable.getEntryAt(idx));
						thisStr = rtss.toString();
					} catch (NumberFormatException ex) {
					}
					break;
				case NUMBER:
					String n = value.toString();
					if (this.formatString != null) {
						thisStr = formatter.formatRawCellContents(Double.parseDouble(n), this.formatIndex, this.formatString);
					} else {
						thisStr = n;
					}
					break;
				default:
					break;
				}
				handler.handleXlsxCell(currRowIndex, currColumnIndex, thisStr);
			}
		}

		@Override
		public void characters(char[] ch, int start, int length) throws SAXException {
			if (vIsOpen) {
				value.append(ch, start, length);
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

		/**
		 * 解码
		 * 
		 * @param inhex
		 * @return
		 */
		protected int decode(String inhex) {
			int column = -1;
			for (int i = 0; i < inhex.length(); ++i) {
				int c = inhex.charAt(i);
				column = (column + 1) * 26 + c - 'A';
			}
			return column;
		}

	}

}
