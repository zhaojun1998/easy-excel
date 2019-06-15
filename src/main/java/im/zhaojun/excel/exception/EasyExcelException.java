package im.zhaojun.excel.exception;

public class EasyExcelException extends RuntimeException {

    private static final long serialVersionUID = 171796064829296469L;

    public EasyExcelException() {
        super();
    }

    public EasyExcelException(String message) {
        super(message);
    }

    public EasyExcelException(String message, Throwable cause) {
        super(message, cause);
    }

    public EasyExcelException(Throwable cause) {
        super(cause);
    }

    protected EasyExcelException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
