package cn.emay.excel.common.schema.annotation;

import java.lang.annotation.*;

/**
 * 列定义<br/>
 * <p>
 * 支持String,Long,Integer,Double,Boolean,Date,BigDecimal类型数据的读写
 *
 * @author Frank
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelColumn {

    /**
     * 列序号[从0开始],不能重复<br/>
     * <p>
     * - writer:匹配列，写入数据；<br/>
     * - reader:当SheetSchema.readColumnBy=Index时，以此进行列-字段的读取匹配；<br/>
     */
    int index();

    /**
     * 列名<br/>
     * <p>
     * - writer:当SheetSchema.isWriteTile=true时，写入第一行；（如果为空，则写入字段名）<br/>
     * - reader:不能重复，当SheetSchema.readColumnBy=Title时，以此进行列-字段的读取匹配；<br/>
     */
    String title();

    /**
     * 列的数据转换表达式<br/>
     * <p>
     * - writer : 写入日期时：格式化日期; 写入Double、BigDecimal时：是保留的小数点后数字个数；<br/>
     * - reader : 读取日期时：如果是String写入，则根据此表达式进行格式化读取；<br/>
     * - reader : 读取Double、BigDecimal时，是保留的小数点后数字个数；<br/>
     */
    String express() default "";

}
