package cn.emay.excel.write;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import cn.emay.excel.common.ExcelVersion;
import cn.emay.excel.common.schema.base.SheetSchema;
import cn.emay.excel.utils.ExcelWriteUtils;
import cn.emay.excel.write.data.SchemaSheetDataGetter;
import cn.emay.excel.write.data.SheetDataGetter;
import cn.emay.excel.write.data.WriteData;
import cn.emay.excel.write.writer.SheetWriter;
import cn.emay.excel.write.writer.impl.SchemaSheetWriter;

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
	 * 把Excel写入文件【根据后缀（.xls,.xlsx）自动适配】<br/>
	 * dataClass实现了@ExcelSheet注解,其字段实现了@ExcelColumn注解
	 * 
	 * @param excelPath
	 *            Excel写入的全路径
	 * @param datas
	 *            数据集,按照顺序写入
	 */
	public static void write(String excelPath, List<?>... datas) {
		write(null, null, excelPath, datas);
	}

	/**
	 * 把Excel写入输出流<br/>
	 * dataClass实现了@ExcelSheet注解,其字段实现了@ExcelColumn注解
	 * 
	 * @param os
	 *            输出流
	 * @param version
	 *            版本
	 * @param datas
	 *            写入的数据集,按照顺序写入
	 */
	public static void write(OutputStream os, ExcelVersion version, List<?>... datas) {
		write(os, version, null, datas);
	}

	/**
	 * 整合方法
	 * 
	 * @param os
	 *            输出流
	 * @param version
	 *            版本
	 * @param excelPath
	 *            Excel写入的全路径
	 * @param datas
	 *            写入的数据集,按照顺序写入
	 */
	@SuppressWarnings({ "unchecked", "rawtypes" })
	private static void write(OutputStream os, ExcelVersion version, String excelPath, List<?>... datas) {
		if (datas == null) {
			throw new IllegalArgumentException("datas is null");
		}
		if (datas.length == 0) {
			throw new IllegalArgumentException("datas is empty");
		}
		SheetWriter[] handlers = new SheetWriter[datas.length];
		int cacheNumber = DEFAULT_CACHE_NUM;
		for (int i = 0; i < datas.length; i++) {
			List<?> list = datas[i];
			Class<?> dataClass = list.get(0).getClass();
			SheetDataGetter<?> dataWriter = new ListSchemaSheetDataGetter(list, dataClass);
			SheetSchema sheetSchema = new SheetSchema(dataClass);
			handlers[i] = new SchemaSheetWriter(sheetSchema, dataWriter);
			cacheNumber = sheetSchema.getSheetWriteSchemaParams().getCacheNumber() > cacheNumber ? sheetSchema.getSheetWriteSchemaParams().getCacheNumber() : cacheNumber;
		}
		if (excelPath != null) {
			write(excelPath, cacheNumber, handlers);
		} else {
			write(os, version, cacheNumber, handlers);
		}
	}

	/**
	 * 把Excel写入文件【根据后缀（.xls,.xlsx）自动适配】<br/>
	 * dataClass实现了@ExcelSheet注解,其字段实现了@ExcelColumn注解
	 * 
	 * @param excelPath
	 *            Excel写入的全路径
	 * @param datas
	 *            数据集,按照顺序写入
	 */
	public static void write(String excelPath, SheetDataGetter<?>... datas) {
		write(null, null, excelPath, datas);
	}

	/**
	 * 把Excel写入输出流<br/>
	 * dataClass实现了@ExcelSheet注解,其字段实现了@ExcelColumn注解
	 * 
	 * @param os
	 *            输出流
	 * @param version
	 *            版本
	 * @param datas
	 *            写入的数据集,按照顺序写入
	 */
	public static void write(OutputStream os, ExcelVersion version, SheetDataGetter<?>... datas) {
		write(os, version, null, datas);
	}

	/**
	 * 整合方法
	 * 
	 * @param os
	 *            输出流
	 * @param version
	 *            版本
	 * @param excelPath
	 *            Excel写入的全路径
	 * @param datas
	 *            写入的数据集,按照顺序写入
	 */
	@SuppressWarnings({ "unchecked", "rawtypes" })
	private static void write(OutputStream os, ExcelVersion version, String excelPath, SheetDataGetter<?>... datas) {
		if (datas == null) {
			throw new IllegalArgumentException("datas is null");
		}
		if (datas.length == 0) {
			throw new IllegalArgumentException("datas is empty");
		}
		SheetWriter[] handlers = new SheetWriter[datas.length];
		int cacheNumber = DEFAULT_CACHE_NUM;
		for (int i = 0; i < datas.length; i++) {
			SheetDataGetter<?> dataWriter = datas[i];
			SheetSchema sheetSchema = new SheetSchema(datas[i].getDataClass());
			handlers[i] = new SchemaSheetWriter(sheetSchema, dataWriter);
			cacheNumber = sheetSchema.getSheetWriteSchemaParams().getCacheNumber() > cacheNumber ? sheetSchema.getSheetWriteSchemaParams().getCacheNumber() : cacheNumber;
		}
		if (excelPath != null) {
			write(excelPath, cacheNumber, handlers);
		} else {
			write(os, version, cacheNumber, handlers);
		}
	}

	/**
	 * 把Excel写入文件【根据后缀（.xls,.xlsx）自动适配】
	 * 
	 * @param excelPath
	 *            Excel写入的全路径
	 * @param dataWriterByCustomSchema
	 *            数据来源集合，每一个数据来源写入一个sheet
	 */
	public static void write(String excelPath, SchemaSheetDataGetter<?>... datas) {
		write(null, null, excelPath, datas);
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
	public static void write(OutputStream os, ExcelVersion version, SchemaSheetDataGetter<?>... datas) {
		write(os, version, null, datas);
	}

	/**
	 * 整合方法
	 * 
	 * @param os
	 *            输出流
	 * @param version
	 *            版本
	 * @param excelPath
	 *            Excel写入的全路径
	 * @param datas
	 *            写入的数据集,按照顺序写入
	 */
	@SuppressWarnings({ "unchecked", "rawtypes" })
	private static void write(OutputStream os, ExcelVersion version, String excelPath, SchemaSheetDataGetter<?>... datas) {
		if (datas == null) {
			throw new IllegalArgumentException("datas is null");
		}
		if (datas.length == 0) {
			throw new IllegalArgumentException("datas is empty");
		}
		SheetWriter[] handlers = new SheetWriter[datas.length];
		int cacheNumber = DEFAULT_CACHE_NUM;
		for (int i = 0; i < datas.length; i++) {
			SheetDataGetter<?> dataWriter = datas[i];
			SheetSchema sheetSchema = datas[i].getSheetSchema();
			handlers[i] = new SchemaSheetWriter(sheetSchema, dataWriter);
			cacheNumber = sheetSchema.getSheetWriteSchemaParams().getCacheNumber() > cacheNumber ? sheetSchema.getSheetWriteSchemaParams().getCacheNumber() : cacheNumber;
		}
		if (excelPath != null) {
			write(excelPath, cacheNumber, handlers);
		} else {
			write(os, version, cacheNumber, handlers);
		}
	}

	/**
	 * 把Excel写入文件【根据后缀（.xls,.xlsx）自动适配】
	 * 
	 * @param excelPath
	 *            Excel写入的全路径
	 * @param handlers
	 *            Execl写入处理器集合[按照顺序处理Sheet,SheetWriteHandler实例不要重用]
	 */
	public static void write(String excelPath, SheetWriter... handlers) {
		write(excelPath, DEFAULT_CACHE_NUM, handlers);
	}

	/**
	 * 把Excel写入文件【根据后缀（.xls,.xlsx）自动适配】
	 * 
	 * @param excelPath
	 *            Excel写入的全路径
	 * @param cacheNumber
	 *            在内存中的缓存数据行数(XLSX适用)【小于100直接使用先写入内存再全部刷到磁盘的方式;大于100则采用当内存中超过CacheNumber条后，刷新到磁盘的方式】
	 * @param handlers
	 *            Execl写入处理器集合[按照顺序处理Sheet,SheetWriteHandler实例不要重用]
	 */
	public static void write(String excelPath, int cacheNumber, SheetWriter... handlers) {
		if (excelPath == null) {
			throw new IllegalArgumentException("excelPath is null");
		}
		if (handlers == null) {
			throw new IllegalArgumentException("handlers is null");
		}
		if (handlers.length == 0) {
			throw new IllegalArgumentException("handlers is empty");
		}
		ExcelVersion version = null;
		if (excelPath.endsWith(ExcelVersion.XLS.getSuffix())) {
			version = ExcelVersion.XLS;
		} else if (excelPath.endsWith(ExcelVersion.XLSX.getSuffix())) {
			version = ExcelVersion.XLSX;
		} else {
			throw new IllegalArgumentException("is not excel file  : " + excelPath);
		}
		File file = new File(excelPath);
		if (file.exists()) {
			throw new IllegalArgumentException("excelPath[" + excelPath + "]  is exists");
		}
		boolean error = false;
		FileOutputStream fos = null;
		File parent = file.getParentFile();
		try {
			if (!parent.exists()) {
				parent.mkdirs();
			}
			fos = new FileOutputStream(excelPath);
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
	 * 把Excel写入文件【根据后缀（.xls,.xlsx）自动适配】
	 * 
	 * @param excelPath
	 *            Excel写入的全路径
	 * @param cacheNumber
	 *            在内存中的缓存数据行数(XLSX适用)【小于100直接使用先写入内存再全部刷到磁盘的方式;大于100则采用当内存中超过CacheNumber条后，刷新到磁盘的方式】
	 * @param handlers
	 *            Execl写入处理器集合[按照顺序处理Sheet,SheetWriteHandler实例不要重用]
	 */
	public static void writeBigData(String excelPath, int cacheNumber, SheetWriter... handlers) {
		if (excelPath == null) {
			throw new IllegalArgumentException("excelPath is null");
		}
		if (handlers == null) {
			throw new IllegalArgumentException("handlers is null");
		}
		if (handlers.length == 0) {
			throw new IllegalArgumentException("handlers is empty");
		}
		ExcelVersion version = null;
		if (excelPath.endsWith(ExcelVersion.XLS.getSuffix())) {
			version = ExcelVersion.XLS;
		} else if (excelPath.endsWith(ExcelVersion.XLSX.getSuffix())) {
			version = ExcelVersion.XLSX;
		} else {
			throw new IllegalArgumentException("is not excel file  : " + excelPath);
		}
		File file = new File(excelPath);
		if (file.exists()) {
			throw new IllegalArgumentException("excelPath[" + excelPath + "]  is exists");
		}
		boolean error = false;
		FileOutputStream fos = null;
		File parent = file.getParentFile();
		try {
			if (!parent.exists()) {
				parent.mkdirs();
			}
			fos = new FileOutputStream(excelPath);
			writeBigData(fos, version, cacheNumber, handlers);
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
	public static void write(OutputStream os, ExcelVersion version, int cacheNumber, SheetWriter... handlers) {
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
					if (SXSSFWorkbook.class.isAssignableFrom(workbook.getClass())) {
						((SXSSFWorkbook) workbook).dispose();
					}
				} catch (IOException e) {
					throw new IllegalArgumentException(e);
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
	 * @param cacheNumber
	 *            在内存中的缓存数据行数(XLSX适用)【小于100直接使用先写入内存再全部刷到磁盘的方式;大于100则采用当内存中超过CacheNumber条后，刷新到磁盘的方式】
	 * @param handlers
	 *            Execl写入处理器集合[按照顺序处理Sheet,SheetWriteHandler实例不要重用]
	 */
	public static void writeBigData(OutputStream os, ExcelVersion version, int cacheNumber, SheetWriter... handlers) {
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
			writeBigData(workbook, handlers);
			workbook.write(os);
			os.flush();
		} catch (IOException e) {
			throw new IllegalArgumentException(e);
		} finally {
			if (workbook != null) {
				try {
					workbook.close();
					if (SXSSFWorkbook.class.isAssignableFrom(workbook.getClass())) {
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
			if (SXSSFSheet.class.isAssignableFrom(sheet.getClass())) {
				((SXSSFSheet) sheet).trackAllColumnsForAutoSizing();
			}
			write(index, sheet, handler);
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
	public static void writeBigData(Workbook workbook, SheetWriter... handlers) {
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
			Row row = sheet.createRow(rowIndex);
			handler.beginRow(rowIndex);
			for (int columnIndex = 0; columnIndex <= handler.getMaxColumnIndex(); columnIndex++) {
				Cell cell = row.createCell(columnIndex);
				cell.setCellStyle(cell.getSheet().getWorkbook().createCellStyle());
				handler.writeCell(cell, rowIndex, columnIndex);
			}
			handler.endRow(rowIndex);
			rowIndex++;
		}
		handler.end(sheetIndex);
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
	public static void writeBigData(int sheetIndex, Sheet sheet, SheetWriter handler) {
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
		CellStyle createCellStyle = sheet.getWorkbook().createCellStyle();
		while (handler.hasRow(rowIndex)) {
			Row row = sheet.createRow(rowIndex);
			handler.beginRow(rowIndex);
			for (int columnIndex = 0; columnIndex <= handler.getMaxColumnIndex(); columnIndex++) {
				Cell cell = row.createCell(columnIndex);
				cell.setCellStyle(createCellStyle);
				handler.writeCell(cell, rowIndex, columnIndex);
			}
			handler.endRow(rowIndex);
			rowIndex++;
		}
		handler.end(sheetIndex);
	}

	/**
	 * 往已经存在的Excel文件中写入数据<br/>
	 * 写完后重新生成一份文件<br/>
	 * 大数据Excel禁用
	 * 
	 * @param fromExcelPath
	 *            读取的Excel文件
	 * @param toExcelPath
	 *            Excel写入数据后，生成新的Excel文件
	 * @param datas
	 *            数据
	 */
	public static void writeByCoordinate(String fromExcelPath, String toExcelPath, WriteData... datas) {
		if (fromExcelPath == null) {
			throw new IllegalArgumentException("fromExcelPath is null");
		}
		if (toExcelPath == null) {
			throw new IllegalArgumentException("toExcelPath is null");
		}
		if (datas == null) {
			throw new IllegalArgumentException("datas is null");
		}
		ExcelVersion version;
		if (!new File(fromExcelPath).exists()) {
			throw new IllegalArgumentException("fromExcelPath[" + fromExcelPath + "] is not exists");
		}
		if (fromExcelPath.endsWith(ExcelVersion.XLSX.getSuffix())) {
			version = ExcelVersion.XLSX;
		} else if (fromExcelPath.endsWith(ExcelVersion.XLS.getSuffix())) {
			version = ExcelVersion.XLS;
		} else {
			throw new IllegalArgumentException("fromExcelPath[" + fromExcelPath + "] is not excel");
		}
		Map<Integer, Map<Integer, Map<Integer, WriteData>>> dataByCoordinate = new TreeMap<>();
		for (WriteData data : datas) {
			if (data == null) {
				continue;
			}
			if (data.getSheetIndex() < 0) {
				continue;
			}
			if (data.getRowIndex() < 0) {
				continue;
			}
			if (data.getColumnIndex() < 0) {
				continue;
			}
			Map<Integer, Map<Integer, WriteData>> oneSheetMap = dataByCoordinate.get(data.getSheetIndex());
			if (oneSheetMap == null) {
				oneSheetMap = new TreeMap<>();
				dataByCoordinate.put(data.getSheetIndex(), oneSheetMap);
			}
			Map<Integer, WriteData> oneRowMap = oneSheetMap.get(data.getRowIndex());
			if (oneRowMap == null) {
				oneRowMap = new TreeMap<>();
				oneSheetMap.put(data.getRowIndex(), oneRowMap);
			}
			oneRowMap.put(data.getColumnIndex(), data);
		}
		FileInputStream fis = null;
		FileOutputStream fos = null;
		Workbook workbook = null;
		boolean error = false;
		File to = new File(toExcelPath);
		File parent = to.getParentFile();
		try {
			fis = new FileInputStream(fromExcelPath);
			if (ExcelVersion.XLSX.equals(version)) {
				workbook = new XSSFWorkbook(fis);
			} else {
				workbook = new HSSFWorkbook(fis);
			}
			int maxSheetIndex = workbook.getNumberOfSheets() - 1;
			for (Integer sheetIndex : dataByCoordinate.keySet()) {
				Sheet sheet = null;
				if (sheetIndex > maxSheetIndex) {
					sheet = workbook.createSheet();
					maxSheetIndex++;
				} else {
					sheet = workbook.getSheetAt(sheetIndex);
				}
				Map<Integer, Map<Integer, WriteData>> oneSheetMap = dataByCoordinate.get(sheetIndex);
				for (Integer rowIndex : oneSheetMap.keySet()) {
					Row row = sheet.getRow(rowIndex);
					if (row == null) {
						row = sheet.createRow(rowIndex);
					}
					Map<Integer, WriteData> oneRowMap = oneSheetMap.get(rowIndex);
					for (Integer columnIndex : oneRowMap.keySet()) {
						WriteData data = oneRowMap.get(columnIndex);
						Cell cell = row.getCell(columnIndex);
						if (cell == null) {
							cell = row.createCell(columnIndex);
						}
						ExcelWriteUtils.write(cell, data.getData(), data.getExpress());
					}
				}
			}
			if (!parent.exists()) {
				parent.mkdirs();
			}
			fos = new FileOutputStream(to);
			workbook.write(fos);
			fos.flush();
		} catch (IOException e) {
			throw new IllegalArgumentException(e);
		} finally {
			if (workbook != null) {
				try {
					workbook.close();
					if (SXSSFWorkbook.class.isAssignableFrom(workbook.getClass())) {
						((SXSSFWorkbook) workbook).dispose();
					}
				} catch (IOException e) {
					throw new IllegalArgumentException(e);
				}
			}
			if (fis != null) {
				try {
					fis.close();
				} catch (IOException e) {
					throw new IllegalArgumentException(e);
				}
			}
			if (fos != null) {
				try {
					fos.close();
				} catch (IOException e) {
					throw new IllegalArgumentException(e);
				} finally {
					if (error) {
						to.delete();
						parent.delete();
					}
				}
			}
		}
	}
}

/**
 * 基于现有数据的数据获取器
 * 
 * @author Frank
 *
 * @param <D>
 */
class ListSchemaSheetDataGetter<D> implements SheetDataGetter<D> {

	/**
	 * 数据集合
	 */
	private List<D> datas;
	/**
	 * 数据大小
	 */
	private int size;
	/**
	 * 数据Class
	 */
	private Class<D> dataClass;

	/**
	 * 
	 * @param datas
	 *            数据集合
	 * @param dataClass
	 *            数据Class
	 */
	public ListSchemaSheetDataGetter(List<D> datas, Class<D> dataClass) {
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