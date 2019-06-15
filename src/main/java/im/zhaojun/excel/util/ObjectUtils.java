package im.zhaojun.excel.util;

import im.zhaojun.excel.exception.NotSupportTypeException;
import org.joda.time.DateTime;
import org.joda.time.format.DateTimeFormat;

import java.util.Date;

public class ObjectUtils {

    /**
     * 是否是数值类型
     */
    public static boolean isNumeric(String str) {
        for (int i = str.length(); --i >= 0; ) {
            if (!Character.isDigit(str.charAt(i))) {
                return false;
            }
        }
        return true;
    }

    public static Object parseStringToBasicDataType(String value, Class<?> fieldType) {
        if (Byte.class.equals(fieldType) || Byte.TYPE.equals(fieldType)) {
            return Byte.valueOf(value);
        } else if (Boolean.class.equals(fieldType) || Boolean.TYPE.equals(fieldType)) {
            return Boolean.valueOf(value) || "1".equals(value);
        } else if (String.class.equals(fieldType)) {
            return value;
        } else if (Short.class.equals(fieldType) || Short.TYPE.equals(fieldType)) {
            return Short.valueOf(value);
        } else if (Integer.class.equals(fieldType) || Integer.TYPE.equals(fieldType)) {
            return Integer.valueOf(value);
        } else if (Long.class.equals(fieldType) || Long.TYPE.equals(fieldType)) {
            return Long.valueOf(value);
        } else if (Float.class.equals(fieldType) || Float.TYPE.equals(fieldType)) {
            return Float.valueOf(value);
        } else if (Double.class.equals(fieldType) || Double.TYPE.equals(fieldType)) {
            return Double.valueOf(value);
        } else {
            throw new NotSupportTypeException("Illegal data type: " + fieldType);
        }
    }

    public static Date parseDate(String value) {
        return parseDate(value, "");
    }

    public static Date parseDate(String value, String format) {
        if ("".equals(format)) {
            return ObjectUtils.isNumeric(value) ? new Date(Long.valueOf(value)) : DateTime.parse(value).toDate();
        } else {
            return DateTime.parse(value, DateTimeFormat.forPattern(format)).toDate();
        }
    }
}
