package cn.emay.excel.write.data;

import cn.emay.excel.common.schema.base.SheetSchema;

/**
 * 基于定义的表数据获取器
 *
 * @param <D>
 * @author Frank
 */
public interface SchemaSheetDataGetter<D> extends SheetDataGetter<D> {

    /**
     * 表定义
     *
     * @return 表定义
     */
    SheetSchema getSheetSchema();

}
