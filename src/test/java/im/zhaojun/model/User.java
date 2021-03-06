package im.zhaojun.model;

import im.zhaojun.excel.annotation.EasyExcelColumn;
import im.zhaojun.excel.annotation.EasyExcelMapping;
import im.zhaojun.excel.annotation.EasyExcelMappings;
import im.zhaojun.excel.annotation.EasyExcelSheet;

import java.util.Date;

@EasyExcelSheet(startRow = 1, sheetIndex = 1)
public class User {

    @EasyExcelColumn(index = 0)
    private String username;

    @EasyExcelColumn(index = 1)
    private Integer age;

    @EasyExcelMappings({
        @EasyExcelMapping(key = "男", value = "1"),
        @EasyExcelMapping(key = "女", value = "0")
    })
    @EasyExcelColumn(index = 2)
    private Integer sex;

    @EasyExcelColumn(index = 4, name = "创建日期")
    private Date createTime;

    public String getUsername() {
        return username;
    }

    public void setUsername(String username) {
        this.username = username;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public Integer getSex() {
        return sex;
    }

    public void setSex(Integer sex) {
        this.sex = sex;
    }

    public Date getCreateTime() {
        return createTime;
    }

    public void setCreateTime(Date createTime) {
        this.createTime = createTime;
    }

    @Override
    public String toString() {
        return "User{" +
                "username='" + username + '\'' +
                ", age=" + age +
                ", sex=" + sex +
                ", createTime=" + createTime +
                '}';
    }
}
