package im.zhaojun.model;

import im.zhaojun.excel.annotation.EasyExcelProperty;
import im.zhaojun.excel.annotation.EasyExcelSheet;

import java.util.Date;

@EasyExcelSheet(value = "学生名单", headRow = 1)
public class User {

    @EasyExcelProperty(index = 0)
    private String username;

    @EasyExcelProperty(index = 1)
    private int age;

    @EasyExcelProperty(index = 2)
    private Date birthday;

    @EasyExcelProperty(index = 3)
    private String sex;

    @Override
    public String toString() {
        return "User{" +
                "username='" + username + '\'' +
                ", age=" + age +
                ", birthday=" + birthday +
                ", sex='" + sex + '\'' +
                '}';
    }

    public String getUsername() {
        return username;
    }

    public void setUsername(String username) {
        this.username = username;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public Date getBirthday() {
        return birthday;
    }

    public void setBirthday(Date birthday) {
        this.birthday = birthday;
    }

    public String getSex() {
        return sex;
    }

    public void setSex(String sex) {
        this.sex = sex;
    }
}
