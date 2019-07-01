package im.zhaojun.excel.handler;

public abstract class ExcelRowHandler<T> {

    public void before() {}

    public abstract void execute(T t);

    public void doAfterAll() {}

}