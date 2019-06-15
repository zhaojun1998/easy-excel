package im.zhaojun.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface EasyExcelSheet {


    /**
     * 根据名称获取 Sheet 页
     */
    String value() default "";

    /**
     * 根据坐标获取 Sheet 页, 从 0 开始. 如指定了此参数, 则会忽略 {@link #value()} 参数.
     */
    int index() default -1;

    int headRow() default 0;
}
