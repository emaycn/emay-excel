package cn.emay.excel.reader.handler;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;

import cn.emay.excel.common.Person;
import cn.emay.excel.read.ExcelReadHelper;
import cn.emay.excel.read.handler.SheetReadHandler;

/**
 * 
 * @author Frank
 *
 */
public class ReadNormalHandler implements SheetReadHandler {

	private List<Person> list = new ArrayList<>();
	private List<String> titles = new ArrayList<String>();

	private Person curr;

	@Override
	public int getStartReadRowIndex() {
		return 0;
	}

	@Override
	public int getEndReadRowIndex() {
		return -1;
	}

	@Override
	public void begin(int sheetIndex, String sheetName) {

	}

	@Override
	public void beginRow(int rowIndex) {
		curr = new Person();
	}

	@Override
	public void handleXlsxCell(int rowIndex, int columnIndex, int formatIndex, String value) {
		if (rowIndex == 0) {
			titles.add(ExcelReadHelper.readString(value));
		} else {
			switch (columnIndex) {
			case 0:
				curr.setAge(ExcelReadHelper.readInteger(value));
				break;
			case 1:
				curr.setName(ExcelReadHelper.readString(value));
				break;
			case 2:
				curr.setBrith(ExcelReadHelper.readDate(formatIndex, value, "yyyy-MM-dd HH:mm:ss"));
				break;
			case 3:
				curr.setCreateTime(ExcelReadHelper.readLong(value));
				break;
			case 4:
				curr.setScore(ExcelReadHelper.readDouble(value, 2));
				break;
			case 5:
				curr.setHasGlass(ExcelReadHelper.readBoolean(value));
				break;
			case 6:
				curr.setMoney(ExcelReadHelper.readBigDecimal(value, 4));
				break;
			default:
				break;
			}
		}
	}

	@Override
	public void handleXlsCell(int rowIndex, int columnIndex, Cell cell) {
		if (rowIndex == 0) {
			titles.add(ExcelReadHelper.readString(cell));
		} else {
			switch (columnIndex) {
			case 0:
				curr.setAge(ExcelReadHelper.readInteger(cell));
				break;
			case 1:
				curr.setName(ExcelReadHelper.readString(cell));
				break;
			case 2:
				curr.setBrith(ExcelReadHelper.readDate(cell, "yyyy-MM-dd HH:mm:ss"));
				break;
			case 3:
				curr.setCreateTime(ExcelReadHelper.readLong(cell));
				break;
			case 4:
				curr.setScore(ExcelReadHelper.readDouble(cell, 2));
				break;
			case 5:
				curr.setHasGlass(ExcelReadHelper.readBoolean(cell));
				break;
			case 6:
				curr.setMoney(ExcelReadHelper.readBigDecimal(cell, 4));
				break;
			default:
				break;
			}
		}
	}

	@Override
	public void endRow(int rowIndex) {
		if (rowIndex > 0 && curr != null) {
			list.add(curr);
		}
	}

	@Override
	public void end(int sheetIndex, String sheetName) {

	}

	public List<Person> getDatas() {
		return list;
	}

	public List<String> getTitles() {
		return titles;
	}

}
