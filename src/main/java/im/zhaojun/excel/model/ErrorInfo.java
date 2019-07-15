package im.zhaojun.excel.model;

/**
 * @author Zhao Jun
 * 2019/7/15 19:55
 */
public class ErrorInfo {

    private String coordinate;

    private Object value;

    private String name;

    public String getCoordinate() {
        return coordinate;
    }

    public void setCoordinate(String coordinate) {
        this.coordinate = coordinate;
    }

    public Object getValue() {
        return value;
    }

    public void setValue(Object value) {
        this.value = value;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    @Override
    public String toString() {
        return "ErrorInfo{" +
                "coordinate='" + coordinate + '\'' +
                ", value=" + value +
                ", name='" + name + '\'' +
                '}';
    }
}
