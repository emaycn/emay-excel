package cn.emay.excel.support;

import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.List;
import java.util.UUID;

import javax.servlet.http.HttpServletResponse;

import cn.emay.excel.common.ExcelVersion;
import cn.emay.excel.write.ExcelWriter;
import cn.emay.excel.write.handler.SheetWriteHandler;
import cn.emay.excel.write.handler.impl.DataWriter;

/**
 * Excel针对Servlet的支持
 * 
 * @author Frank
 *
 */
public class ExcelServletSupport {

	/**
	 * 导出Excel
	 * 
	 * @param response
	 *            响应
	 * @param excelName
	 *            Excel文件名(不包含.xls/.xlsx后缀)
	 * @param version
	 *            Excel版本
	 * @param handlers
	 *            写Sheet处理器
	 */
	public static void outputWithExcel(HttpServletResponse response, String excelName, ExcelVersion version, SheetWriteHandler... handlers) {
		try {
			if (handlers == null || handlers.length == 0) {
				throw new IllegalArgumentException("excel handlers is empty!");
			}
			checkAndFill(response, excelName, version);
			ExcelWriter.writeExcel(response.getOutputStream(), version, handlers);
		} catch (IOException e) {
			throw new IllegalArgumentException(e);
		}
	}

	/**
	 * 导出Excel
	 * 
	 * @param response
	 *            响应
	 * @param excelName
	 *            Excel文件名(不包含.xls/.xlsx后缀)
	 * @param version
	 *            Excel版本
	 * @param datas
	 *            写Sheet的数据处理器（数据要实现@ExcelSheet、@ExcelColumn注解）
	 */
	public static void outputWithExcel(HttpServletResponse response, String excelName, ExcelVersion version, DataWriter<?>... datas) {
		try {
			if (datas == null || datas.length == 0) {
				throw new IllegalArgumentException("excel datas is empty!");
			}
			checkAndFill(response, excelName, version);
			ExcelWriter.writeExcel(response.getOutputStream(), version, datas);
		} catch (IOException e) {
			throw new IllegalArgumentException(e);
		}
	}

	/**
	 * 导出Excel
	 * 
	 * @param response
	 *            响应
	 * @param excelName
	 *            Excel文件名(不包含.xls/.xlsx后缀)
	 * @param version
	 *            Excel版本
	 * @param datas
	 *            写Sheet的数据（数据要实现@ExcelSheet、@ExcelColumn注解）
	 */
	public static void outputWithExcel(HttpServletResponse response, String excelName, ExcelVersion version, List<?>... datas) {
		try {
			if (datas == null || datas.length == 0) {
				throw new IllegalArgumentException("excel datas is empty!");
			}
			checkAndFill(response, excelName, version);
			ExcelWriter.writeExcel(response.getOutputStream(), version, datas);
		} catch (IOException e) {
			throw new IllegalArgumentException(e);
		}
	}

	/**
	 * 检测及填充数据
	 * 
	 * @param response
	 *            响应
	 * @param excelName
	 *            Excel名字
	 * @param version
	 *            Excel版本
	 * @throws UnsupportedEncodingException
	 */
	public static void checkAndFill(HttpServletResponse response, String excelName, ExcelVersion version) throws UnsupportedEncodingException {
		if (response == null) {
			throw new IllegalArgumentException("response is null!");
		}
		if (excelName == null) {
			excelName = UUID.randomUUID().toString().replace("-", "");
		}
		if (version == null) {
			throw new IllegalArgumentException("excel version is null!");
		}
		// response.setContentType("application/x-download");
		if (ExcelVersion.XLS.equals(version)) {
			response.setContentType("application/vnd.ms-excel");
		} else {
			response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
		}
		excelName = URLEncoder.encode(excelName, "UTF-8");
		response.addHeader("Content-Disposition", "attachment;filename=" + excelName + version.getSuffix());
	}

}
