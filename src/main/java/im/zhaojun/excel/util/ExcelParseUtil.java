package im.zhaojun.excel.util;

import im.zhaojun.excel.annotation.EasyExcelSheet;
import im.zhaojun.excel.annotation.FieldType;
import im.zhaojun.excel.metadata.Sheet;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Date;

public class ExcelParseUtil {

    private static final Logger log = LoggerFactory.getLogger(ExcelParseUtil.class);

    public static Object getDataValue(FieldType fieldType, String value, SharedStringsTable sharedStringsTable) {
        if (null == value || "".equals(value)) {
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
    public static Sheet parseSheet(Class<?> clz) {
        Sheet sheet = new Sheet();

        sheet.setClazz(clz);

        EasyExcelSheet easyExcelSheet = clz.getDeclaredAnnotation(EasyExcelSheet.class);

        if (easyExcelSheet == null) {
            log.debug("未获取到注解, 默认解析第一个 sheet 页");
            log.debug("未获取到注解, 默认从第一行开始读取数据");
            sheet.setSheetNo(1);
            sheet.setStartRow(1);
        } else {
            log.debug("解析第 [{}] 个 sheet 页", easyExcelSheet.sheetIndex());
            log.debug("从第 [{}] 行开始解析数据", easyExcelSheet.sheetIndex());
            sheet.setSheetNo(easyExcelSheet.sheetIndex());
            sheet.setStartRow(easyExcelSheet.startRow());
        }
        return sheet;
    }

    public static Date parseDate(String value, String format) {
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern(format);
        return Date.from(LocalDate.parse(value, formatter).atStartOfDay(ZoneId.systemDefault()).toInstant());
    }

    public static String convertString(Object object) {
        return (String) object;
    }


    public static boolean objIsString(Object object) {
        return object instanceof String;
    }
}