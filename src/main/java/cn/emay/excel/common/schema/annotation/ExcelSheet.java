package cn.emay.excel.common.schema.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 表定义
 * 
 * @author Frank
 *
 */
@Target({ ElementType.TYPE })
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelSheet {

	/*-------------read---------------*/

	/**
	 * 标题行号<br/>
	 * 
	 * - reader:从哪一行读取标题，如果小于0则不读取标题，默认为0。【如果readColumnBy=Title，则不可以小于0】<br/>
	 * 
	 * @return
	 */
	int readTitleRowIndex() default 0;

	/**
	 * 开始读取数据的行号(从0开始)<br/>
	 * 当readColumnBy=Title时，数据行号必须比title号行要大<br/>
	 * 
	 * - reader:从哪一行开始读取，默认index=1<br/>
	 * 
	 * @return
	 */
	int readDataStartRowIndex() default 1;

	/**
	 * 结束读取数据的行号(从0开始)<br/>
	 * 当readColumnBy=Title时，数据行号必须比title号行要大<br/>
	 * 
	 * - reader:读取到哪一行结束，如果小于0则全部读取，默认-1<br/>
	 * 
	 * @return
	 */
	int readDataEndRowIndex() default -1;

	/**
	 * 读取列规则[Index,Title];<br/>
	 * 
	 * - reader:匹配数据根据，默认根据列Index进行数据匹配。如果=Ttile，【则readtitleRowIndex不可以小于0】<br/>
	 * 
	 * @return
	 */
	String readColumnBy() default "Index";

	/*-------------write---------------*/

	/**
	 * 表名<br/>
	 * 
	 * - writer:如果不为空，将表名写入Excel；<br/>
	 * 
	 * @return
	 */
	String writeSheetName() default "";

	/**
	 * 是否写入title,默认true<br/>
	 * 
	 * - writer:是否将每个字段的title写在第一行；<br/>
	 * 
	 * @return
	 */
	boolean isWriteTile() default true;

	/**
	 * 写入缓存条数<br/>
	 * 
	 * -
	 * writer:当写xlsx时，如果writeCacheNumber>=1000，实时刷盘；如果writeCacheNumber<1000，内存构建完成后刷盘；<br/>
	 * 
	 * @return
	 */
	int cacheNumber() default 1000;

	/**
	 * 是否自适应宽度，默认开启<br/>
	 * 
	 * - writer:每一列取最长数据宽度的125%，有微量性能损失；<br/>
	 * 
	 * @return
	 */
	boolean isAutoWidth() default false;

	/**
	 * 表头背景色<br/>
	 * 
	 * - writer:RGB自定义背景色设置，默认全白；<br/>
	 * 
	 * @return
	 */
	int[] titleRgbColor() default { 255, 255, 255 };

	/**
	 * 内容列背景色<br/>
	 * 
	 * - writer:RGB自定义背景色设置，默认全白；<br/>
	 * 
	 * @return
	 */
	int[] contentRgbColor() default { 255, 255, 255 };

	/**
	 * 是否需要单元格边框，默认false<br/>
	 * 
	 * - writer:单元格边框画线；<br/>
	 * 
	 * @return
	 */
	boolean isNeedBorder() default false;

	/**
	 * 是否自动换行，默认是<br/>
	 * 
	 * - writer:单元格自动换行；<br/>
	 * 
	 * @return
	 */
	boolean isAutoWrap() default true;

}
