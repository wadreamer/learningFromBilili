package EasyExcel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.read.metadata.ReadSheet;
import org.junit.Test;

import java.io.File;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * ClassName: Test
 * Description:
 * date: 2020/7/7 9:20
 *
 * @author CFG
 * @since JDK 1.8
 */
public class EasyExcelTest {

    private final static String PATH = "D:/IntelliJ IDEA 2019.3.3/idea_workplace/excel_poi/src/main/resources";

    private List<DemoData> data() {
        List<DemoData> list = new ArrayList<DemoData>();
        for (int i = 0; i < 10; i++) {
            DemoData data = new DemoData();
            data.setString("字符串" + i);
            data.setDate(new Date());
            data.setDoubleData(0.56);
            list.add(data);
        }
        return list;
    }

    @Test
    public void simpleWrite() {
        // 写法1
        String fileName = PATH + "/EasyExcelTest.xlsx";

        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        // 如果这里想使用03 则 传入excelType参数即可
        // write(fileName, 格式类)
        // sheet(表名)
        // doWrite(数据)
        EasyExcel.write(fileName, DemoData.class).sheet("模板").doWrite(data());

    }

    @Test
    public void simpleRead() {
        // 有个很重要的点 DemoDataListener 不能被spring管理，要每次读取excel都要new,然后里面用到spring可以构造方法传进去
        // 写法1：
        String fileName = PATH + "/EasyExcelTest.xlsx";
        // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭
        EasyExcel.read(fileName, DemoData.class, new DemoDataListener()).sheet().doRead();

    }

}


