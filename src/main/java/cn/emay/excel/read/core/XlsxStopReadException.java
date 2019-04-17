package cn.emay.excel.read.core;

/**
 * 
 * SAX读取方式停止异常
 * 
 * @author Frank
 *
 */
public class XlsxStopReadException extends RuntimeException {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	public XlsxStopReadException() {
		super();
	}

	public XlsxStopReadException(String message) {
		super(message);
	}

}