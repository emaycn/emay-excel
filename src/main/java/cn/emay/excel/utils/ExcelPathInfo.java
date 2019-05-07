package cn.emay.excel.utils;

import java.io.InputStream;

import cn.emay.excel.common.ExcelVersion;

/**
 * 路径信息
 * 
 * @author Frank
 *
 */
public class ExcelPathInfo {

	/**
	 * 版本
	 */
	private ExcelVersion version;

	/**
	 * 文件输入流
	 */
	private InputStream inputStream;

	public ExcelPathInfo() {

	}

	public ExcelPathInfo(ExcelVersion version, InputStream inputStream) {
		this.version = version;
		this.setInputStream(inputStream);
	}

	/**
	 * 版本
	 * 
	 * @return
	 */
	public ExcelVersion getVersion() {
		return version;
	}

	public void setVersion(ExcelVersion version) {
		this.version = version;
	}

	/**
	 * 文件输入流
	 * 
	 * @return
	 */
	public InputStream getInputStream() {
		return inputStream;
	}

	public void setInputStream(InputStream inputStream) {
		this.inputStream = inputStream;
	}

}
