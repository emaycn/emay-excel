package cn.emay.excel.read.core;

import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import cn.emay.excel.read.handler.SheetReadHandler;

/**
 * SAX方式读取ExcelSheet页
 * 
 * @author Frank
 *
 */
public class XlsxSheetHandler extends DefaultHandler {

	/**
	 * Sheet读处理器
	 */
	private SheetReadHandler handler;
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

	public XlsxSheetHandler(StylesTable stylesTable, SharedStringsTable sst, int sheetIndex, String sheetName, SheetReadHandler handler) {
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