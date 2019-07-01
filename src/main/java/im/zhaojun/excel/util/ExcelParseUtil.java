package im.zhaojun.excel.util;

import cn.hutool.core.util.StrUtil;
import im.zhaojun.excel.annotation.EasyExcelSheet;
import im.zhaojun.excel.annotation.FieldType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Date;

/**
 * @author Zhao Jun
 * 2019/6/27 21:56
 */
public class ExcelParseUtil {

    private static final Logger log = LoggerFactory.getLogger(ExcelParseUtil.class);

    public static Object getDataValue(FieldType fieldType, String value, SharedStringsTable sharedStringsTable, String numFmtString) {
        if (null == value) {
            return null;
        }
        Object result;
        switch (fieldType) {
            case BOOLEAN:
                result = (value.charAt(0) != '0');
                break;
            case ERROR:
                result = StrUtil.format("\\\"ERROR: {} ", value);
                break;
            case STRING:
                try {
                    final int index = Integer.parseInt(value);
                    result = new XSSFRichTextString(sharedStringsTable.getEntryAt(index)).getString();
                } catch (NumberFormatException e) {
                    result = value;
                }
                break;
            case NUMBER:
                result = getNumberValue(value, numFmtString);
                break;
            case DATE:
                try {
                    result = DateUtil.getJavaDate(Double.parseDouble(value));
                } catch (Exception e) {
                    result = value;
                }
                break;
            default:
                result = value;
                break;
        }
        return result;
    }


    /**
     * 获取数字类型值
     *
     * @param value 值
     * @param numFmtString 格式
     * @return 数字，可以是Double、Long
     * @since 4.1.0
     */
    private static Number getNumberValue(String value, String numFmtString) {
        if(StrUtil.isBlank(value)) {
            return null;
        }
        double numValue = Double.parseDouble(value);
        // 普通数字
        if (null != numFmtString && numFmtString.indexOf(StrUtil.C_DOT) < 0) {
            final long longPart = (long) numValue;
            if (longPart == numValue) {
                // 对于无小数部分的数字类型，转为Long
                return longPart;
            }
        }
        return numValue;
    }

    /**
     * 根据 Class 类获取 Excel Sheet.
     */
    public static int parseSheet(Class<?> clz) {
        EasyExcelSheet easyExcelSheet = clz.getDeclaredAnnotation(EasyExcelSheet.class);

        if (easyExcelSheet == null) {
            log.debug("未获取到注解, 默认解析第一个 sheet 页");
            return 1;
        }

        return easyExcelSheet.sheetIndex();
    }

    /**
     * 根据 Class 类获取 Excel Sheet.
     */
    public static int parseStartRow(Class<?> clz) {
        EasyExcelSheet easyExcelSheet = clz.getDeclaredAnnotation(EasyExcelSheet.class);

        if (easyExcelSheet == null) {
            log.debug("未获取到注解, 默认解析第一个 sheet 页");
            return 1;
        }

        return easyExcelSheet.startRow();
    }


    public static Date parseDate(String value, String format) {
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern(format);
        return Date.from(LocalDate.parse(value, formatter).atStartOfDay(ZoneId.systemDefault()).toInstant());
    }

    public static Number convertNumber(Object object) {
        return (Number) object;
    }

    public static String convertString(Object object) {
        return (String) object;
    }

    public static boolean objIsNumber(Object object) {
        return object instanceof Number;
    }

    public static boolean objIsString(Object object) {
        return object instanceof String;
    }

    public static boolean objIsBoolean(Object object) {
        return object instanceof Boolean;
    }

    public static boolean objIsDate(Object object) {
        return object instanceof Date;
    }

}