package cn.emay.excel.common.schema.base;

/**
 * 表定义的写参数集
 * 
 * @author Frank
 *
 */
public class SheetWriteSchemaParams {

	/**
	 * 表名<br/>
	 * 
	 * - writer:如果不为空，将表名写入Excel；<br/>
	 */
	private String writeSheetName = "";

	/**
	 * 是否写入title,默认true<br/>
	 * 
	 * - writer:是否将每个字段的title写在第一行；<br/>
	 */
	private boolean isWriteTile = true;

	/**
	 * 写入缓存条数，默认1000<br/>
	 * 
	 * -
	 * writer:当写xlsx时，如果writeCacheNumber>=1000，实时刷盘；如果writeCacheNumber<1000，内存构建完成后刷盘；<br/>
	 */
	private int cacheNumber = 1000;

	/**
	 * 是否自适应宽度，默认开启<br/>
	 * 
	 * - writer:每一列取最长数据宽度的125%，有微量性能损失；<br/>
	 */
	private boolean isAutoWidth = true;

	/**
	 * 表头背景色<br/>
	 * 
	 * - writer:RGB自定义背景色设置，默认全白；<br/>
	 */
	private int[] titleRgbColor = { 255, 255, 255 };

	/**
	 * 内容列背景色<br/>
	 * 
	 * - writer:RGB自定义背景色设置，默认全白；<br/>
	 */
	private int[] contentRgbColor = { 255, 255, 255 };

	/**
	 * 是否需要单元格边框，默认false<br/>
	 * 
	 * - writer:单元格边框画线；<br/>
	 */
	private boolean isNeedBorder = false;

	/**
	 * 是否自动换行，默认是<br/>
	 * 
	 * - writer:单元格自动换行；<br/>
	 */
	private boolean isAutoWrap = true;

	public String getWriteSheetName() {
		return writeSheetName;
	}

	public void setWriteSheetName(String writeSheetName) {
		this.writeSheetName = writeSheetName;
	}

	public boolean isWriteTile() {
		return isWriteTile;
	}

	public void setWriteTile(boolean isWriteTile) {
		this.isWriteTile = isWriteTile;
	}

	public int getCacheNumber() {
		return cacheNumber;
	}

	public void setCacheNumber(int cacheNumber) {
		this.cacheNumber = cacheNumber;
	}

	public boolean isAutoWidth() {
		return isAutoWidth;
	}

	public void setAutoWidth(boolean isAutoWidth) {
		this.isAutoWidth = isAutoWidth;
	}

	public int[] getTitleRgbColor() {
		return titleRgbColor;
	}

	public void setTitleRgbColor(int[] titleRgbColor) {
		this.titleRgbColor = titleRgbColor;
	}

	public int[] getContentRgbColor() {
		return contentRgbColor;
	}

	public void setContentRgbColor(int[] contentRgbColor) {
		this.contentRgbColor = contentRgbColor;
	}

	public boolean isNeedBorder() {
		return isNeedBorder;
	}

	public void setNeedBorder(boolean isNeedBorder) {
		this.isNeedBorder = isNeedBorder;
	}

	public boolean isAutoWrap() {
		return isAutoWrap;
	}

	public void setAutoWrap(boolean isAutoWrap) {
		this.isAutoWrap = isAutoWrap;
	}

}