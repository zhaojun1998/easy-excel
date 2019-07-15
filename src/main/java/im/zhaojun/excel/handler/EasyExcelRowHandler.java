package im.zhaojun.excel.handler;

import im.zhaojun.excel.context.EasyExcelContext;

public abstract class EasyExcelRowHandler<T> {

    public void before() {}

    public abstract void execute(T t, EasyExcelContext context);

    public void doAfterAll(EasyExcelContext context) {}

}