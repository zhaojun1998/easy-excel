package im.zhaojun.excel.annotation;

public enum FieldType {

    STRING("s"),

    NUMBER(""),

    DATE("m/d/yy"),

    BOOLEAN("b"),

    EMPTY(""),

    ERROR("e");

    private String name;

    FieldType(String name) {
        this.name = name;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }


    /**
     * 类型字符串转为枚举
     *
     * @param name 类型字符串
     * @return 类型枚举
     */
    public static FieldType of(String name) {
        if (null == name) {
            //默认数字
            return NUMBER;
        }

        if (BOOLEAN.name.equals(name)) {
            return BOOLEAN;
        } else if (ERROR.name.equals(name)) {
            return ERROR;
        } else if (STRING.name.equals(name)) {
            return STRING;
        } else {
            return EMPTY;
        }
    }
}