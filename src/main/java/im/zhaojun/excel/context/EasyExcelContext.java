package im.zhaojun.excel.context;

import im.zhaojun.excel.handler.EasyExcelRowHandler;
import im.zhaojun.excel.metadata.Sheet;
import im.zhaojun.excel.model.ErrorInfo;

import java.io.InputStream;
import java.util.List;

public class EasyExcelContext {

    private InputStream inputStream;

    private Sheet currentSheet;

    private EasyExcelRowHandler handler;

    private Integer currentRowNum;

    private Integer totalCount;

    private List<ErrorInfo> errorInfoList;

    private boolean fastFail;

    public EasyExcelContext(InputStream inputStream, Sheet currentSheet, EasyExcelRowHandler handler, boolean fastFail) {
        this.inputStream = inputStream;
        this.currentSheet = currentSheet;
        this.handler = handler;
        this.fastFail = fastFail;
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

    public List<ErrorInfo> getErrorInfoList() {
        return errorInfoList;
    }

    public void setErrorInfoList(List<ErrorInfo> errorInfoList) {
        this.errorInfoList = errorInfoList;
    }

    public boolean getFastFail() {
        return fastFail;
    }

    public void setFastFail(boolean fastFail) {
        this.fastFail = fastFail;
    }
}
