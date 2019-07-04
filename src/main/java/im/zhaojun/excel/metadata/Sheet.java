package im.zhaojun.excel.metadata;

public class Sheet {

    /**
     * Sheet 名称.
     */
    private String sheetName;

    /**
     * Sheet 序号, 从 1 开始
     */
    private int sheetNo;

    /**
     * 开始行数, 从 1 开始
     */
    private int startRow;

    /**
     * 要转换成的对应的 Class 类
     */
    private Class<?> clazz;

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public int getSheetNo() {
        return sheetNo;
    }

    public void setSheetNo(int sheetNo) {
        this.sheetNo = sheetNo;
    }

    public Class<?> getClazz() {
        return clazz;
    }

    public void setClazz(Class<?> clazz) {
        this.clazz = clazz;
    }

    public int getStartRow() {
        return startRow;
    }

    public void setStartRow(int startRow) {
        this.startRow = startRow;
    }
}
