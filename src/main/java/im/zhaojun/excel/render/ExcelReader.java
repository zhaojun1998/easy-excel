package im.zhaojun.excel.render;

import im.zhaojun.excel.analysis.XlsxSaxAnalyser;
import im.zhaojun.excel.context.EasyExcelContext;
import im.zhaojun.excel.handler.EasyExcelRowHandler;
import im.zhaojun.excel.metadata.Sheet;
import im.zhaojun.excel.util.ExcelParseUtil;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;

import java.io.IOException;
import java.io.InputStream;

public class ExcelReader {

    public static void read(InputStream inputStream, EasyExcelRowHandler handler, Class<?> clz) {
        read(inputStream, handler, clz, true);
    }

    public static void read(InputStream inputStream, EasyExcelRowHandler handler, Class<?> clz, boolean fastFail) {
        Sheet sheet = ExcelParseUtil.parseSheet(clz);
        EasyExcelContext easyExcelContext = new EasyExcelContext(inputStream, sheet, handler, fastFail);
        try {
            new XlsxSaxAnalyser(easyExcelContext).execute();
        } catch (IOException | OpenXML4JException e) {
            e.printStackTrace();
        }
    }


}