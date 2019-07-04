package im.zhaojun.excel.context;

import im.zhaojun.excel.handler.EasyExcelRowHandler;
import im.zhaojun.excel.metadata.Sheet;

import java.io.InputStream;

public class EasyExcelContext {

    private InputStream inputStream;

    private Sheet currentSheet;

    private EasyExcelRowHandler handler;

    private Integer currentRowNum;

    private Integer totalCount;

    public EasyExcelContext(InputStream inputStream, Sheet currentSheet, EasyExcelRowHandler handler) {
        this.inputStream = inputStream;
        this.currentSheet = currentSheet;
        this.handler = handler;
    }

    public Sheet getCurrentSheet() {
        return currentSheet;
    }

    public void setCurrentSheet(Sheet currentSheet) {
        this.currentSheet = currentSheet;
    }

    public EasyExcelRowHandler getHandler() {
        return handler;
    }

    public void setHandler(EasyExcelRowHandler handler) {
        this.handler = handler;
    }

    public Integer getCurrentRowNum() {
        return currentRowNum;
    }

    public void setCurrentRowNum(Integer currentRowNum) {
        this.currentRowNum = currentRowNum;
    }

    public Integer getTotalCount() {
        return totalCount;
    }

    public void setTotalCount(Integer totalCount) {
        this.totalCount = totalCount;
    }

    public InputStream getInputStream() {
        return inputStream;
    }

    public void setInputStream(InputStream inputStream) {
        this.inputStream = inputStream;
    }
}
