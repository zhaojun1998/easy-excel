package im.zhaojun.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface EasyExcelSheet {

    /**
     * 根据坐标获取 Sheet 页, 坐标从 0 开始, 默认为 0.
     */
    int sheetIndex() default 1;

    int startRow() default 0;
}