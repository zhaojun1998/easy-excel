package im.zhaojun;

import im.zhaojun.excel.context.EasyExcelContext;
import im.zhaojun.excel.handler.EasyExcelRowHandler;
import im.zhaojun.excel.render.ExcelReader;
import im.zhaojun.excel.util.FileUtil;
import im.zhaojun.model.User;
import org.junit.Test;

import java.io.InputStream;

public class EasyExcelParseTest {

    @Test
    public void testEasyExcel() {
        InputStream inputStream = FileUtil.getResourcesFileInputStream("user.xlsx");

        ExcelReader.read(inputStream, new EasyExcelRowHandler<User>() {
            int i = 0;

            @Override
            public void execute(User user, EasyExcelContext context) {
                i++;
                System.out.println(user);
            }

            @Override
            public void doAfterAll(EasyExcelContext context) {
                System.out.println("count:" + i);
                System.out.println(context.getErrorInfoList());
            }
        }, User.class, false);
    }
}
