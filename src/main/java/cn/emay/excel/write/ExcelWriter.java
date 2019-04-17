package cn.emay.excel.write;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import cn.emay.excel.common.ExcelSheet;
import cn.emay.excel.common.ExcelVersion;
import cn.emay.excel.write.handler.SheetWriteHandler;
import cn.emay.excel.write.handler.impl.DataWriter;
import cn.emay.excel.write.handler.impl.SheetWriteBySchemaHandler;

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
	private final static int DEFAULT_CACHE_NUM = 100;

	/**
	 * 把Excel写入文件【根据后缀（.xls,.xlsx）自动适配】
	 * 
	 * @param excelabsolutePath
	 *            Excel写入的全路径
	 * @param datas
	 *            数据集,按照顺序写入
	 */
	@SuppressWarnings({ "unchecked", "rawtypes" })
	public static void writeExcel(String excelAbsolutePath, List<?>... datas) {
		if (datas == null || datas.length == 0) {
			throw new IllegalArgumentException("datas is null or empty");
		}
		NormalData<?>[] datawrite = new NormalData<?>[datas.length];
		for (int i = 0; i < datas.length; i++) {
			List<?> data = datas[i];
			datawrite[i] = new NormalData(data);
		}
		writeExcel(excelAbsolutePath, datawrite);
	}

	/**
	 * 把Excel写入文件【根据后缀（.xls,.xlsx）自动适配】
	 * 
	 * @param excelabsolutePath
	 *            Excel写入的全路径
	 * @param datas
	 *            数据集处理器,按照顺序写入[处理器实例不可复用]
	 */
	public static void writeExcel(String excelAbsolutePath, DataWriter<?>... datas) {
		if (datas == null || datas.length == 0) {
			throw new IllegalArgumentException("datas is null or empty");
		}
		SheetWriteHandler[] handlers = new SheetWriteHandler[datas.length];
		int cacheNumber = 0;
		for (int i = 0; i < datas.length; i++) {
			DataWriter<?> data = datas[i];
			if (data.getShemaClass() == null) {
				throw new IllegalArgumentException("schemaClass is null");
			}
			if (!data.getShemaClass().isAnnotationPresent(ExcelSheet.class)) {
				throw new IllegalArgumentException("schemaClass[" + data.getShemaClass().getName() + "] is not has Annotation : " + ExcelSheet.class.getName());
			}
			ExcelSheet schema = data.getShemaClass().getAnnotation(ExcelSheet.class);
			@SuppressWarnings({ "unchecked", "rawtypes" })
			SheetWriteHandler handler = new SheetWriteBySchemaHandler(data);
			handlers[i] = handler;
			cacheNumber = cacheNumber > schema.cacheNumber() ? cacheNumber : schema.cacheNumber();
		}
		writeExcel(excelAbsolutePath, cacheNumber, handlers);
	}

	/**
	 * 把Excel写入文件【根据后缀（.xls,.xlsx）自动适配】
	 * 
	 * @param excelabsolutePath
	 *            Excel写入的全路径
	 * @param handlers
	 *            Execl写入处理器集合[按照顺序处理Sheet,SheetWriteHandler实例不要重用]
	 */
	public static void writeExcel(String excelAbsolutePath, SheetWriteHandler... handlers) {
		writeExcel(excelAbsolutePath, 1000, handlers);
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
	public static void writeExcel(String excelAbsolutePath, int cacheNumber, SheetWriteHandler... handlers) {
		boolean error = false;
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
		if (file.getParentFile().exists()) {
			if (!file.getParentFile().isDirectory()) {
				throw new IllegalArgumentException("excelAbsolutePath's parent file[" + file.getParentFile().getAbsolutePath() + "]  is not a dir");
			}
		} else {
			file.getParentFile().mkdirs();
		}
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(excelAbsolutePath);
			writeExcel(fos, version, cacheNumber, handlers);
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
						new File(excelAbsolutePath).delete();
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
	 * @param datas
	 *            写入的数据集[多sheet],按照顺序写入
	 */
	@SuppressWarnings({ "unchecked", "rawtypes" })
	public static void writeExcel(OutputStream os, ExcelVersion version, List<?>... datas) {
		if (datas == null || datas.length == 0) {
			throw new IllegalArgumentException("datas is null");
		}
		NormalData<?>[] datawrite = new NormalData<?>[datas.length];
		for (int i = 0; i < datas.length; i++) {
			List<?> data = datas[i];
			datawrite[i] = new NormalData(data);
		}
		writeExcel(os, version, datawrite);
	}

	/**
	 * 把Excel写入输出流<br/>
	 * 
	 * @param os
	 *            输出流
	 * @param version
	 *            版本
	 * @param datas
	 *            写入的数据集处理器,按照顺序写入[处理器实例不可复用]
	 */
	public static void writeExcel(OutputStream os, ExcelVersion version, DataWriter<?>... datas) {
		if (datas == null || datas.length == 0) {
			throw new IllegalArgumentException("datas is null");
		}
		int cacheNumber = 0;
		SheetWriteHandler[] handlers = new SheetWriteHandler[datas.length];
		for (int i = 0; i < datas.length; i++) {
			DataWriter<?> data = datas[i];
			if (data.getShemaClass() == null) {
				throw new IllegalArgumentException("schemaClass is null");
			}
			if (!data.getShemaClass().isAnnotationPresent(ExcelSheet.class)) {
				throw new IllegalArgumentException("schemaClass[" + data.getShemaClass().getName() + "] is not has Annotation : " + ExcelSheet.class.getName());
			}
			ExcelSheet schema = data.getShemaClass().getAnnotation(ExcelSheet.class);
			@SuppressWarnings({ "unchecked", "rawtypes" })
			SheetWriteHandler handler = new SheetWriteBySchemaHandler(data);
			handlers[i] = handler;
			cacheNumber = cacheNumber > schema.cacheNumber() ? cacheNumber : schema.cacheNumber();
		}
		writeExcel(os, version, cacheNumber, handlers);
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
	public static void writeExcel(OutputStream os, ExcelVersion version, SheetWriteHandler... handlers) {
		writeExcel(os, version, 1000, handlers);
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
	public static void writeExcel(OutputStream os, ExcelVersion version, int cacheNumber, SheetWriteHandler... handlers) {
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
		try {
			workbook = createWorkbook(version, cacheNumber);
			writeData(workbook, handlers);
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
			try {
				os.close();
			} catch (IOException e) {
				throw new IllegalArgumentException(e);
			}
		}
	}

	/**
	 * 构建WorkBook
	 * 
	 * @param version
	 *            版本
	 * @param cacheNumber
	 *            缓存条数
	 * @return
	 */
	private static Workbook createWorkbook(ExcelVersion version, int cacheNumber) {
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
		return workbook;
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
	private static Workbook writeData(Workbook workbook, SheetWriteHandler... handlers) {
		for (int i = 0; i < handlers.length; i++) {
			SheetWriteHandler handler = handlers[i];
			Sheet sheet = null;
			if (handler == null) {
				sheet = workbook.createSheet();
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
			handler.begin(i);
			int rowIndex = 0;
			while (handler.hasRow(rowIndex)) {
				Row row = sheet.createRow(rowIndex);
				handler.beginRow(rowIndex);
				for (int columnIndex = 0; columnIndex <= handler.getMaxColumnIndex(); columnIndex++) {
					Cell cell = row.createCell(columnIndex);
					CellStyle style = cell.getSheet().getWorkbook().createCellStyle();
					cell.setCellStyle(style);
					handler.writeCell(cell, rowIndex, columnIndex);
				}
				handler.endRow(rowIndex);
				rowIndex++;
			}
			handler.end(i);
		}
		return workbook;
	}

}

class NormalData<R extends Object> implements DataWriter<R> {

	private List<R> datas;
	int size;

	public NormalData(List<R> datas) {
		this.datas = datas;
		size = datas.size();
	}

	@Override
	public R getData(int rowIndex) {
		return datas.get(rowIndex);
	}

	@Override
	public boolean hasData(int rowIndex) {
		return rowIndex < size;
	}

	@SuppressWarnings("unchecked")
	@Override
	public Class<R> getShemaClass() {
		return (Class<R>) datas.get(0).getClass();
	}

}
