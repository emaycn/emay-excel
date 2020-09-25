package cn.emay.excel.utils;

import cn.emay.excel.common.ExcelVersion;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 一些工具
 *
 * @author Frank
 */
public class ExcelUtils {

    /**
     * 获取int类型的表达式
     *
     * @return int
     */
    public static int parserExpressToInt(String express) {
        int num = -1;
        if (express != null && !"".equalsIgnoreCase(express.trim())) {
            try {
                return Integer.parseInt(express);
            } catch (Exception ignored) {
            }
        }
        return num;
    }

    /**
     * 把字符串转成日期
     *
     * @param dateStr 日期字符串
     * @param format  日期格式
     * @return 日期
     */
    public static Date parseDate(String dateStr, String format) {
        Date date = null;
        try {
            SimpleDateFormat sdf = new SimpleDateFormat(format);
            date = sdf.parse(dateStr);
        } catch (Exception ignored) {
        }
        return date;
    }

    /**
     * 新建一个数据实例
     *
     * @return 数据
     */
    public static <D> D newData(Class<D> dataClass) {
        try {
            return dataClass.newInstance();
        } catch (InstantiationException | IllegalAccessException e) {
            throw new IllegalArgumentException(dataClass.getName() + " can't be new Instance", e);
        }
    }

    /**
     * 从路径中解析出版本及输入流
     *
     * @param excelPath Excel路径
     * @return 版本
     */
    public static ExcelVersion parserPath(String excelPath) {
        ExcelVersion version;
        if (excelPath == null) {
            throw new IllegalArgumentException("excelPath is null");
        }
        if (!new File(excelPath).exists()) {
            throw new IllegalArgumentException("excelPath[" + excelPath + "] is not exists");
        }
        if (excelPath.endsWith(ExcelVersion.XLSX.getSuffix())) {
            version = ExcelVersion.XLSX;
        } else if (excelPath.endsWith(ExcelVersion.XLS.getSuffix())) {
            version = ExcelVersion.XLS;
        } else {
            throw new IllegalArgumentException("excelPath[" + excelPath + "] is not excel");
        }
        return version;
    }

}
