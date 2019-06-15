package im.zhaojun.excel;

import im.zhaojun.excel.annotation.EasyExcelProperty;
import im.zhaojun.excel.annotation.EasyExcelSheet;
import im.zhaojun.excel.exception.SheetNotExistException;
import im.zhaojun.excel.util.ObjectUtils;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Field;
import java.text.DecimalFormat;
import java.util.*;

public class EasyExcelParse {

    private static final Logger log = LoggerFactory.getLogger(EasyExcelParse.class);

    public static <T> List<T> parseFromFile(Class<T> clz, File file) {
        try (Workbook workbook = WorkbookFactory.create(file)) {
            return parseFromWorkbook(clz, workbook);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    public static <T> List<T> parseFromWorkbook(Class<T> clz, Workbook workbook) {
        if (workbook == null) throw new NullPointerException("WorkBook is empty");
        Sheet sheet = parseSheet(clz, workbook);

        Row headRow = parseHeadRow(clz, sheet);

        Map<Integer, Field> fieldMap = getFieldMap(clz);

        EasyExcelSheet easyExcelSheet = clz.getDeclaredAnnotation(EasyExcelSheet.class);

        int headRowNum = easyExcelSheet == null ? 0 : easyExcelSheet.headRow();

        List<T> list = new ArrayList<>();

        for (int i = headRowNum; i < sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            try {
                T t = clz.newInstance();

                for (Cell cell : row) {
                    String value = getCellValueAsString(cell);
                    invoke(t, fieldMap.get(cell.getColumnIndex()), value);
                }
                list.add(t);
            } catch (InstantiationException | IllegalAccessException e) {
                e.printStackTrace();
            }
        }
        return list;
    }

    private static <T> void invoke(T t, Field field, String value) throws IllegalAccessException {
        field.setAccessible(true);
        field.set(t, parseValueWithFieldType(field, value));
    }

    private static Object parseValueWithFieldType(Field field, String value) {
        Class<?> type = field.getType();

        EasyExcelProperty excelProperty = field.getDeclaredAnnotation(EasyExcelProperty.class);

        String format = excelProperty.format();

        // 如果是日期类型, 或字符串类型, 但标注了格式化日志的字段, 则尝试转换成日期格式.
        if (Date.class.equals(type) || (String.class.equals(type) && !"".equals(format))) {
            return ObjectUtils.parseDate(value, format);
        }

        return ObjectUtils.parseStringToBasicDataType(value, type);
    }


    private static String getCellValueAsString(Cell cell) {
        if (null == cell) {
            return "";
        }
        if (cell.getCellType() == CellType.NUMERIC) {
            if (DateUtil.isCellDateFormatted(cell)) {
                return String.valueOf(cell.getDateCellValue().getTime());
            }
            DecimalFormat df = new DecimalFormat();
            return df.format(cell.getNumericCellValue());
        }
        return cell.toString();
    }



    private static <T> Map<Integer, Field> getFieldMap(Class<T> clz) {
        Field[] fields = clz.getDeclaredFields();

        Map<Integer, Field> fieldMap = new HashMap<>();

        for (Field field : fields) {
            EasyExcelProperty easyExcelProperty = field.getAnnotation(EasyExcelProperty.class);
            if (easyExcelProperty != null) {
                fieldMap.put(easyExcelProperty.index(), field);
            }
        }
        return fieldMap;
    }

    /**
     * 获取头信息
     */
    private static <T> Row parseHeadRow(Class<T> clz, Sheet sheet) {
        EasyExcelSheet easyExcelSheet = clz.getAnnotation(EasyExcelSheet.class);

        if (easyExcelSheet == null) {
            log.debug("默认取第一行为 head");
            return sheet.getRow(0);
        }

        int headRow = easyExcelSheet.headRow();

        log.debug("取第 {} 行为 head", headRow);
        return sheet.getRow(headRow);
    }


    /**
     * 根据 Class 类获取 Excel Sheet.
     *
     * @param clz 要转换的 Class 类
     * @return Sheet 页
     */
    private static <T> Sheet parseSheet(Class<T> clz, Workbook workbook) {
        EasyExcelSheet easyExcelSheet = clz.getAnnotation(EasyExcelSheet.class);

        if (easyExcelSheet == null) {
            log.debug("默认解析第一个 sheet 页");
            return workbook.getSheetAt(0);
        }

        int index = easyExcelSheet.index();
        String value = easyExcelSheet.value();

        if ("".equals(value) && index == -1) {
            log.debug("默认解析第一个 sheet 页");
            return workbook.getSheetAt(0);
        }

        try {
            if (index != -1) {
                log.debug("根据 index 解析 sheet 页: {}", index);
                return workbook.getSheetAt(index);
            }
        } catch (IllegalArgumentException e) {
            throw new SheetNotExistException("不存在的 Sheet Index 页 : " + index);
        }

        Sheet sheet = workbook.getSheet(value);

        log.debug("根据名称解析 sheet 页: {}", value);

        if (sheet == null) {
            throw new SheetNotExistException("不存在名为 [" + value + "] 的 Sheet 页.");
        }
        return sheet;
    }
}
