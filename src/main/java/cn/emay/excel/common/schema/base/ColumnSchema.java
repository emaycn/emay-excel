package cn.emay.excel.common.schema.base;

/**
 * 列定义<br/>
 * 
 * 支持String,Long,Integer,Double,Boolean,Date,BigDecimal类型数据的读写
 * 
 * @author Frank
 *
 */
public class ColumnSchema {

	/**
	 * 列序号[从0开始],不能重复<br/>
	 * 
	 * - writer:匹配列，写入数据；<br/>
	 * - reader:当SheetSchema.readColumnBy=Index时，以此进行列-字段的读取匹配；<br/>
	 */
	private int index;

	/**
	 * 列名<br/>
	 * 
	 * - writer:当SheetSchema.isWriteTile=true时，写入第一行；（如果为空，则写入字段名）<br/>
	 * - reader:不能重复，当SheetSchema.readColumnBy=Title时，以此进行列-字段的读取匹配；<br/>
	 */
	private String title;

	/**
	 * 列的数据转换表达式<br/>
	 * 
	 * - writer : 写入日期时：格式化日期; 写入Double、BigDecimal时：是保留的小数点后数字个数；<br/>
	 * - reader : 读取日期时：如果是String写入，则根据此表达式进行格式化读取；<br/>
	 * - reader : 读取Double、BigDecimal时，是保留的小数点后数字个数；<br/>
	 */
	private String express;

	/**
	 * 
	 */
	public ColumnSchema() {

	}

	/**
	 * 
	 * @param index
	 *            列序号
	 * @param title
	 *            列名
	 * @param express
	 *            列的数据转换表达式
	 */
	public ColumnSchema(int index, String title, String express) {
		this.index = index;
		this.title = title;
		this.express = express;
	}

	public String getTitle() {
		return title;
	}

	public void setTitle(String title) {
		this.title = title;
	}

	public String getExpress() {
		return express;
	}

	public void setExpress(String express) {
		this.express = express;
	}

	public int getIndex() {
		return index;
	}

	public void setIndex(int index) {
		this.index = index;
	}

}
