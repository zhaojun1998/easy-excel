package im.zhaojun.excel.annotation;


import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface EasyExcelColumn {

    String[] value() default {""};

    int index();

    /**
     * 用于标注单元格中的日期格式.
     * 当单元格格式为 日期 时, 不需要此参数也能自动识别.
     * 当单元格格式为 常规 时, 需要指定日期格式.
     */
    String format() default "yyyy-MM-dd HH:mm:ss";

    String name() default "";
}