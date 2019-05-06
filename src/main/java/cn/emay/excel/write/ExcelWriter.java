package cn.emay.excel.write;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import cn.emay.excel.common.ExcelVersion;
import cn.emay.excel.common.schema.base.SheetSchema;
import cn.emay.excel.write.data.SchemaSheetDataGetter;
import cn.emay.excel.write.writer.SheetWriter;
import cn.emay.excel.write.writer.impl.SchemaSheetWriter;
import cn.emay.excel.write.data.SheetDataGetter;

/**
 * Excel写
 * 
 * @author Frank
 *
 */
public class ExcelWriter {

	/**
	 * 默认缓存数据数量
	 */
	public final static int DEFAULT_CACHE_NUM = 1000;

	/**
	 * 把Excel写入文件【根据后缀（.xls,.xlsx）自动适配】
	 * 
	 * @param excelabsolutePath
	 *            Excel写入的全路径
	 * @param datas
	 *            数据集,按照顺序写入
	 */
	public static <D> void writeFirst(String excelAbsolutePath, List<D> datas) {
		if (datas == null || datas.size() == 0) {
			throw new IllegalArgumentException("datas is null");
		}
		@SuppressWarnings("unchecked")
		Class<D> dataClass = (Class<D>) datas.get(0).getClass();
		writeFirst(excelAbsolutePath, new ListSchemaSheetData<D>(datas, dataClass));
	}

	/**
	 * 把Excel写入输出流<br/>
	 * 
	 * @param os
	 *            输出流
	 * @param version
	 *            版本
	 * @param datas
	 *            写入的数据集
	 */
	public static <D> void writeFirst(OutputStream os, ExcelVersion version, List<D> datas) {
		if (datas == null || datas.size() == 0) {
			throw new IllegalArgumentException("datas is null");
		}
		@SuppressWarnings("unchecked")
		Class<D> dataClass = (Class<D>) datas.get(0).getClass();
		writeFirst(os, version, new ListSchemaSheetData<D>(datas, dataClass));
	}

	/**
	 * 把Excel写入文件【根据后缀（.xls,.xlsx）自动适配】
	 * 
	 * @param excelabsolutePath
	 *            Excel写入的全路径
	 * @param datas
	 *            数据集,按照顺序写入
	 */
	public static <D> void writeFirst(String excelAbsolutePath, SheetDataGetter<D> datas) {
		if (datas == null) {
			throw new IllegalArgumentException("datas is null");
		}
		SheetSchema<D> schema = new SheetSchema<D>(datas.getDataClass());
		SheetWriter handler = new SchemaSheetWriter<D>(schema, datas);
		write(excelAbsolutePath, schema.getSheetSchemaParams().getCacheNumber(), handler);
	}

	/**
	 * 把Excel写入输出流<br/>
	 * 
	 * @param os
	 *            输出流
	 * @param version
	 *            版本
	 * @param datas
	 *            写入的数据集
	 */
	public static <D> void writeFirst(OutputStream os, ExcelVersion version, SheetDataGetter<D> datas) {
		if (datas == null) {
			throw new IllegalArgumentException("datas is null");
		}
		SheetSchema<D> schema = new SheetSchema<D>(datas.getDataClass());
		SheetWriter handler = new SchemaSheetWriter<D>(schema, datas);
		write(os, version, schema.getSheetSchemaParams().getCacheNumber(), handler);
	}

	/**
	 * 把Excel写入文件【根据后缀（.xls,.xlsx）自动适配】
	 * 
	 * @param excelabsolutePath
	 *            Excel写入的全路径
	 * @param dataWriterByCustomSchema
	 *            数据来源集合，每一个数据来源写入一个sheet
	 */
	public static <D> void writeFirst(String excelAbsolutePath, SheetSchema<D> sheetSchema, SheetDataGetter<D> dataWriter) {
		SheetWriter handler = new SchemaSheetWriter<D>(sheetSchema, dataWriter);
		write(excelAbsolutePath, sheetSchema.getSheetSchemaParams().getCacheNumber(), handler);
	}

	/**
	 * 把Excel写入输出流<br/>
	 * 
	 * @param os
	 *            输出流
	 * @param version
	 *            版本
	 * @param dataWriterByCustomSchema
	 *            数据来源集合，每一个数据来源写入一个sheet
	 */
	public static <D> void writeFirst(OutputStream os, ExcelVersion version, SheetSchema<D> sheetSchema, SheetDataGetter<D> dataWriter) {
		SheetWriter handler = new SchemaSheetWriter<D>(sheetSchema, dataWriter);
		write(os, version, sheetSchema.getSheetSchemaParams().getCacheNumber(), handler);
	}

	/**
	 * 把Excel写入文件【根据后缀（.xls,.xlsx）自动适配】
	 * 
	 * @param excelabsolutePath
	 *            Excel写入的全路径
	 * @param dataWriterByCustomSchema
	 *            数据来源集合，每一个数据来源写入一个sheet
	 */
	@SuppressWarnings({ "unchecked", "rawtypes" })
	public static <D> void write(String excelAbsolutePath, SchemaSheetDataGetter... datas) {
		SheetWriter[] handlers = new SheetWriter[datas.length];
		int cacheNumber = DEFAULT_CACHE_NUM;
		for (int i = 0; i < datas.length; i++) {
			SheetDataGetter<?> dataWriter = datas[i];
			SheetSchema<?> sheetSchema = datas[i].getSheetSchema();
			handlers[i] = new SchemaSheetWriter(sheetSchema, dataWriter);
			cacheNumber = sheetSchema.getSheetSchemaParams().getCacheNumber() > cacheNumber ? sheetSchema.getSheetSchemaParams().getCacheNumber() : cacheNumber;
		}
		write(excelAbsolutePath, cacheNumber, handlers);
	}

	/**
	 * 把Excel写入输出流<br/>
	 * 
	 * @param os
	 *            输出流
	 * @param version
	 *            版本
	 * @param dataWriterByCustomSchema
	 *            数据来源集合，每一个数据来源写入一个sheet
	 */
	@SuppressWarnings({ "unchecked", "rawtypes" })
	public static void write(OutputStream os, ExcelVersion version, SchemaSheetDataGetter... datas) {
		SheetWriter[] handlers = new SheetWriter[datas.length];
		int cacheNumber = DEFAULT_CACHE_NUM;
		for (int i = 0; i < datas.length; i++) {
			SheetDataGetter<?> dataWriter = datas[i];
			SheetSchema<?> sheetSchema = datas[i].getSheetSchema();
			handlers[i] = new SchemaSheetWriter(sheetSchema, dataWriter);
			cacheNumber = sheetSchema.getSheetSchemaParams().getCacheNumber() > cacheNumber ? sheetSchema.getSheetSchemaParams().getCacheNumber() : cacheNumber;
		}
		write(os, version, cacheNumber, handlers);
	}

	/**
	 * 把Excel写入文件【根据后缀（.xls,.xlsx）自动适配】
	 * 
	 * @param excelabsolutePath
	 *            Excel写入的全路径
	 * @param handlers
	 *            Execl写入处理器集合[按照顺序处理Sheet,SheetWriteHandler实例不要重用]
	 */
	public static void write(String excelAbsolutePath, SheetWriter... handlers) {
		write(excelAbsolutePath, DEFAULT_CACHE_NUM, handlers);
	}

	/**
	 * 把Excel写入文件【根据后缀（.xls,.xlsx）自动适配】
	 * 
	 * @param excelabsolutePath
	 *            Excel写入的全路径
	 * @param cacheNumber
	 *            在内存中的缓存数据行数(XLSX适用)【小于100直接使用先写入内存再全部刷到磁盘的方式;大于100则采用当内存中超过CacheNumber条后，刷新到磁盘的方式】
	 * @param handlers
	 *            Execl写入处理器集合[按照顺序处理Sheet,SheetWriteHandler实例不要重用]
	 */
	public static void write(String excelAbsolutePath, int cacheNumber, SheetWriter... handlers) {
		if (excelAbsolutePath == null) {
			throw new IllegalArgumentException("excelAbsolutePath is null");
		}
		ExcelVersion version = null;
		if (excelAbsolutePath.endsWith(ExcelVersion.XLS.getSuffix())) {
			version = ExcelVersion.XLS;
		} else if (excelAbsolutePath.endsWith(ExcelVersion.XLSX.getSuffix())) {
			version = ExcelVersion.XLSX;
		} else {
			throw new IllegalArgumentException("is not excel file  : " + excelAbsolutePath);
		}
		File file = new File(excelAbsolutePath);
		if (file.exists()) {
			throw new IllegalArgumentException("excelAbsolutePath[" + excelAbsolutePath + "]  is exists");
		}
		boolean error = false;
		FileOutputStream fos = null;
		File parent = file.getParentFile();
		try {
			if (!parent.exists()) {
				parent.mkdirs();
			}
			fos = new FileOutputStream(excelAbsolutePath);
			write(fos, version, cacheNumber, handlers);
		} catch (Exception e) {
			error = true;
			throw new IllegalArgumentException(e);
		} finally {
			if (fos != null) {
				try {
					fos.close();
				} catch (IOException e) {
					throw new IllegalArgumentException(e);
				} finally {
					if (error) {
						file.delete();
						parent.delete();
					}
				}
			}
		}
	}

	/**
	 * 把Excel写入输出流<br/>
	 * 
	 * @param os
	 *            输出流
	 * @param version
	 *            版本
	 * @param handlers
	 *            Execl写入处理器集合[按照顺序处理Sheet,SheetWriteHandler实例不要重用]
	 */
	public static void write(OutputStream os, ExcelVersion version, SheetWriter... handlers) {
		write(os, version, DEFAULT_CACHE_NUM, handlers);
	}

	/**
	 * 把Excel写入输出流<br/>
	 * 
	 * @param os
	 *            输出流
	 * @param version
	 *            版本
	 * @param cacheNumber
	 *            在内存中的缓存数据行数(XLSX适用)【小于100直接使用先写入内存再全部刷到磁盘的方式;大于100则采用当内存中超过CacheNumber条后，刷新到磁盘的方式】
	 * @param handlers
	 *            Execl写入处理器集合[按照顺序处理Sheet,SheetWriteHandler实例不要重用]
	 */
	private static void write(OutputStream os, ExcelVersion version, int cacheNumber, SheetWriter... handlers) {
		if (os == null) {
			throw new IllegalArgumentException("OutputStream is null");
		}
		if (version == null) {
			throw new IllegalArgumentException("ExcelVersion is null");
		}
		if (handlers == null) {
			throw new IllegalArgumentException("handlers is null");
		}
		if (handlers.length == 0) {
			throw new IllegalArgumentException("handlers is empty");
		}
		Workbook workbook = null;
		switch (version) {
		case XLS:
			workbook = new HSSFWorkbook();
			break;
		case XLSX:
			if (cacheNumber >= DEFAULT_CACHE_NUM) {
				workbook = new SXSSFWorkbook(cacheNumber);
			} else {
				workbook = new XSSFWorkbook();
			}
			break;
		default:
			throw new IllegalArgumentException("version is error");
		}
		try {
			write(workbook, handlers);
			workbook.write(os);
			os.flush();
		} catch (IOException e) {
			throw new IllegalArgumentException(e);
		} finally {
			if (workbook != null) {
				try {
					workbook.close();
					if (workbook.getClass().isAssignableFrom(SXSSFWorkbook.class)) {
						((SXSSFWorkbook) workbook).dispose();
					}
				} catch (IOException e) {
					throw new IllegalArgumentException(e);
				}
			}
		}
	}

	/**
	 * 往Workbook写入数据
	 * 
	 * @param workbook
	 *            工作簿
	 * @param handlers
	 *            Execl写入处理器集合[按照顺序处理Sheet,SheetWriteHandler实例不要重用]
	 * @return
	 */
	public static void write(Workbook workbook, SheetWriter... handlers) {
		if (workbook == null) {
			throw new IllegalArgumentException("workbook is null");
		}
		if (handlers == null || handlers.length == 0) {
			throw new IllegalArgumentException("handlers is null or empty");
		}
		for (int index = 0; index < handlers.length; index++) {
			SheetWriter handler = handlers[index];
			Sheet sheet = null;
			if (handler == null) {
				workbook.createSheet();
				continue;
			}
			if (handler.getSheetName() != null && !"".equals(handler.getSheetName())) {
				sheet = workbook.createSheet(handler.getSheetName());
			} else {
				sheet = workbook.createSheet();
			}
			if (sheet.getClass().isAssignableFrom(SXSSFSheet.class)) {
				((SXSSFSheet) sheet).trackAllColumnsForAutoSizing();
			}
			write(index, sheet, handler);
		}
	}

	/**
	 * 写入Sheet
	 * 
	 * @param sheetIndex
	 *            sheet 序号
	 * @param sheet
	 *            表
	 * @param handler
	 *            处理器
	 */
	public static void write(int sheetIndex, Sheet sheet, SheetWriter handler) {
		if (sheetIndex < 0) {
			throw new IllegalArgumentException("sheetIndex must bigger than -1");
		}
		if (sheet == null) {
			throw new IllegalArgumentException("sheet is null");
		}
		if (handler == null) {
			throw new IllegalArgumentException("handler is null");
		}
		handler.begin(sheetIndex);
		int rowIndex = 0;
		while (handler.hasRow(rowIndex)) {
			Row row = sheet.getRow(rowIndex);
			if (row == null) {
				row = sheet.createRow(rowIndex);
			}
			handler.beginRow(rowIndex);
			for (int columnIndex = 0; columnIndex <= handler.getMaxColumnIndex(); columnIndex++) {
				Cell cell = row.getCell(columnIndex);
				if (cell == null) {
					cell = row.createCell(columnIndex);
				}
				cell.setCellStyle(cell.getSheet().getWorkbook().createCellStyle());
				handler.writeCell(cell, rowIndex, columnIndex);
			}
			handler.endRow(rowIndex);
			rowIndex++;
		}
		handler.end(sheetIndex);
	}

}

class ListSchemaSheetData<D> implements SheetDataGetter<D> {

	private List<D> datas;
	private int size;
	private Class<D> dataClass;

	public ListSchemaSheetData(List<D> datas, Class<D> dataClass) {
		this.datas = datas;
		size = datas.size();
		this.dataClass = dataClass;
	}

	@Override
	public D getData(int rowIndex) {
		return datas.get(rowIndex);
	}

	@Override
	public boolean hasData(int rowIndex) {
		return rowIndex < size;
	}

	@Override
	public Class<D> getDataClass() {
		return dataClass;
	}

}