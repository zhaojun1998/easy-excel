package im.zhaojun.excel.util;

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

    public static Object getDataValue(FieldType fieldType, String value, SharedStringsTable sharedStringsTable) {
        if (null == value) {
            return null;
        }
        switch (fieldType) {
            case STRING:
                int idx = Integer.parseInt(value);
                return new XSSFRichTextString(sharedStringsTable.getEntryAt(idx)).getString();
            case DATE:
                return DateUtil.getJavaDate(Double.parseDouble(value));
            case BOOLEAN:
                return (value.charAt(0) == '0' ? "FALSE" : "TRUE");
            case ERROR:
                return "\\\"ERROR: " +  value;
            default:
                return value;
        }
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