package cn.emay.excel.writer.handler;

import java.util.List;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;

import cn.emay.excel.common.Person;
import cn.emay.excel.write.ExcelWriterHelper;
import cn.emay.excel.write.writer.SheetWriter;

/**
 * 
 * @author Frank
 *
 */
public class WriteNormalHandler implements SheetWriter {

	private List<Person> datas;
	private List<String> titles;

	private Person curr;

	public WriteNormalHandler(List<String> titles, List<Person> datas) {
		this.datas = datas;
		this.titles = titles;
	}

	@Override
	public String getSheetName() {
		return "personData";
	}

	@Override
	public boolean hasRow(int rowIndex) {
		return rowIndex - 1 < datas.size();
	}

	@Override
	public int getMaxColumnIndex() {
		return titles.size() - 1;
	}

	@Override
	public void begin(int sheetIndex) {

	}

	@Override
	public void beginRow(int rowIndex) {
		if (rowIndex > 0) {
			curr = datas.get(rowIndex - 1);
		}
	}

	@Override
	public void writeCell(Cell cell, int rowIndex, int columnIndex) {
		CellStyle style = cell.getCellStyle();
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		cell.setCellStyle(style);
		if (rowIndex == 0) {
			ExcelWriterHelper.writeString(cell, titles.get(columnIndex));
		} else {
			switch (columnIndex) {
			case 0:
				ExcelWriterHelper.writeInt(cell, curr.getAge());
				break;
			case 1:
				ExcelWriterHelper.writeString(cell, curr.getName());
				break;
			case 2:
				ExcelWriterHelper.writeDate(cell, curr.getBrith(), "yyyy-MM-dd HH:mm:ss");
				break;
			case 3:
				ExcelWriterHelper.writeLong(cell, curr.getCreateTime());
				break;
			case 4:
				ExcelWriterHelper.writeDouble(cell, curr.getScore(), 2);
				break;
			case 5:
				ExcelWriterHelper.writeBoolean(cell, curr.getHasGlass());
				break;
			case 6:
				ExcelWriterHelper.writeBigDecimal(cell, curr.getMoney(), 4);
				break;
			default:
				break;
			}
		}
	}

	@Override
	public void endRow(int rowIndex) {
		curr = null;
	}

	@Override
	public void end(int sheetIndex) {
		
	}

}
