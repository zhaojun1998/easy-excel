package im.zhaojun;

import im.zhaojun.excel.EasyExcelParse;
import im.zhaojun.model.User;
import org.junit.Test;

import java.io.File;
import java.util.List;

public class EasyExcelParseTest {

    @Test
    public void testParseFromFile() {
        List<User> users = EasyExcelParse.parseFromFile(User.class, new File("C:\\Users\\87301\\Desktop\\666.xlsx"));
        System.out.println(users);
    }

}
