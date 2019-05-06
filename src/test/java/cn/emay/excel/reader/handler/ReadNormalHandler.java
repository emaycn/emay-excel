package cn.emay.excel.reader.handler;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;

import cn.emay.excel.common.Person;
import cn.emay.excel.read.reader.SheetReader;
import cn.emay.excel.utils.ExcelReadUtils;

/**
 * 
 * @author Frank
 *
 */
public class ReadNormalHandler implements SheetReader {

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
			titles.add(ExcelReadUtils.readString(value));
		} else {
			switch (columnIndex) {
			case 0:
				curr.setAge(ExcelReadUtils.readInteger(value));
				break;
			case 1:
				curr.setName(ExcelReadUtils.readString(value));
				break;
			case 2:
				curr.setBrith(ExcelReadUtils.readDate(formatIndex, value, "yyyy-MM-dd HH:mm:ss"));
				break;
			case 3:
				curr.setCreateTime(ExcelReadUtils.readLong(value));
				break;
			case 4:
				curr.setScore(ExcelReadUtils.readDouble(value, 2));
				break;
			case 5:
				curr.setHasGlass(ExcelReadUtils.readBoolean(value));
				break;
			case 6:
				curr.setMoney(ExcelReadUtils.readBigDecimal(value, 4));
				break;
			default:
				break;
			}
		}
	}

	@Override
	public void handleXlsCell(int rowIndex, int columnIndex, Cell cell) {
		if (rowIndex == 0) {
			titles.add(ExcelReadUtils.readString(cell));
		} else {
			switch (columnIndex) {
			case 0:
				curr.setAge(ExcelReadUtils.readInteger(cell));
				break;
			case 1:
				curr.setName(ExcelReadUtils.readString(cell));
				break;
			case 2:
				curr.setBrith(ExcelReadUtils.readDate(cell, "yyyy-MM-dd HH:mm:ss"));
				break;
			case 3:
				curr.setCreateTime(ExcelReadUtils.readLong(cell));
				break;
			case 4:
				curr.setScore(ExcelReadUtils.readDouble(cell, 2));
				break;
			case 5:
				curr.setHasGlass(ExcelReadUtils.readBoolean(cell));
				break;
			case 6:
				curr.setMoney(ExcelReadUtils.readBigDecimal(cell, 4));
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
