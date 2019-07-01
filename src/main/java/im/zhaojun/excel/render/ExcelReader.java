package im.zhaojun.excel.render;

import im.zhaojun.excel.analysis.XlsxSaxAnalyser;
import im.zhaojun.excel.handler.ExcelRowHandler;
import im.zhaojun.excel.util.ExcelParseUtil;

import java.io.InputStream;

public class ExcelReader {

    public static void read(InputStream in, ExcelRowHandler handler, Class<?> clz) {
        new XlsxSaxAnalyser(handler, clz).processOneSheet(in, ExcelParseUtil.parseSheet(clz));
    }

}