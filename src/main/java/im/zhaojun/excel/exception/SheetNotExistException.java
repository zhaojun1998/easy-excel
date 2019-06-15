package im.zhaojun.excel.exception;

public class SheetNotExistException extends EasyExcelException {

    private static final long serialVersionUID = 4401962211756394849L;

    public SheetNotExistException() {
        super();
    }

    public SheetNotExistException(String message) {
        super(message);
    }
}
