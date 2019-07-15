package im.zhaojun.excel;

import java.util.List;

/**
 * @author qjwyss
 * @date 2018/12/19
 * @description 讀取excel資料委託介面
 */
public interface ExcelReadDataDelegated {

    /**
     * 每獲取一條記錄，即寫資料
     * 在flume裡每獲取一條記錄即寫，而不必快取起來，可以大大減少記憶體的消耗，這裡主要是針對flume讀取大資料量excel來說的
     *
     * @param sheetIndex    sheet位置
     * @param totalRowCount 該sheet總行數
     * @param curRow        行號
     * @param cellList      行資料
     */
    public abstract void readExcelDate(int sheetIndex, int totalRowCount, int curRow, List<String> cellList);

}