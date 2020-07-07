package poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * ClassName: ExcelWriteTest
 * Description:
 * date: 2020/7/6 14:47
 *
 * @author CFG
 * @since JDK 1.8
 */
public class ExcelWriteTest {

    private final static String PATH = "D:/IntelliJ IDEA 2019.3.3/idea_workplace/excel_poi/src/main/resources";

    // 03 和 07 版本的区别是 文件后缀名，03 ---> xls，07 ---> xlsx

    @Test
    public void testWriteVersion03() throws IOException {
        // 1. 创建一个工作簿 03 版本的
        Workbook workbook = new HSSFWorkbook();
        // 2. 创建一个工作表
        Sheet sheet = workbook.createSheet("03版本的excel测试");

        // 3. 创建第一行
        Row row1 = sheet.createRow(0);
        // 4. 创建单元格，并写入数据
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("统计人");

        Cell cell12 = row1.createCell(1);
        cell12.setCellValue("wadreamer");

        Row row2 = sheet.createRow(1);

        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");

        Cell cell22 = row2.createCell(1);
        cell22.setCellValue(new DateTime().toString("yyyy-MM-dd HH:mm:ss"));

        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "/测试03版本的Excel.xls");

        workbook.write(fileOutputStream);

        fileOutputStream.close();

        System.out.println("测试03版本的Excel.xls ---> 成功输出：");
    }

    @Test
    public void testWriteVersion07() throws IOException {
        // 1. 创建一个工作簿 07 版本的
        Workbook workbook = new XSSFWorkbook();
        // 2. 创建一个工作表
        Sheet sheet = workbook.createSheet("07版本的excel测试");

        // 3. 创建第一行
        Row row1 = sheet.createRow(0);
        // 4. 创建单元格，并写入数据
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("统计人");

        Cell cell12 = row1.createCell(1);
        cell12.setCellValue("wadreamer");

        Row row2 = sheet.createRow(1);

        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");

        Cell cell22 = row2.createCell(1);
        cell22.setCellValue(new DateTime().toString("yyyy-MM-dd HH:mm:ss"));

        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "/测试07版本的Excel.xlsx");

        workbook.write(fileOutputStream);

        fileOutputStream.close();

        System.out.println("测试07版本的Excel.xlsx ---> 成功输出：");
    }

    /*
     * @Author wadreamer
     * @Description //TODO 最多只能处理 65536 行，否则会抛出异常
     *                 过程中写入缓存，不操作磁盘，最后一次性写入磁盘，速度快
     * @Date 15:40 2020/7/6
     * @Param []
     * @return void
     **/
    @Test
    public void testWriteVersion03BigData() throws IOException {

        Long begin = System.currentTimeMillis();

        Workbook workbook = new HSSFWorkbook();

        Sheet sheet = workbook.createSheet("插入大批量数据");

        for (int rowNum = 0; rowNum < 65536; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int colNum = 0; colNum < 10; colNum++) {
                Cell cell = row.createCell(colNum);
                cell.setCellValue(colNum + 1);
            }
        }

        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "/测试03版本的大量数据Excel.xls");

        workbook.write(fileOutputStream);

        fileOutputStream.close();

        Long end = System.currentTimeMillis();

        System.out.println("耗时 ---> " + (end - begin) / 1000.0);
    }

    /*
     * @Author wadreamer
     * @Description //TODO 写数据时，速度非常慢，非常耗内存，也会发生内存溢出，如一次插入 100W 条记录
     *                 可以写较大的数据量，插入的数据量不受限，如 20W 条记录
     * @Date 15:40 2020/7/6
     * @Param []
     * @return void
     **/
    @Test
    public void testWriteVersion07BigData() throws IOException{

        Long begin = System.currentTimeMillis();

        Workbook workbook = new XSSFWorkbook();

        Sheet sheet = workbook.createSheet("插入大批量数据");

        for (int rowNum = 0; rowNum < 100000; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int colNum = 0; colNum < 10; colNum++) {
                Cell cell = row.createCell(colNum);
                cell.setCellValue(colNum + 1);
            }
        }

        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "/测试07版本的大量数据Excel.xlsx");

        workbook.write(fileOutputStream);

        fileOutputStream.close();

        Long end = System.currentTimeMillis();

        System.out.println("耗时 ---> " + (end - begin) / 1000.0);
    }

    /*
     * @Author wadreamer
     * @Description //TODO 过程中会产生临时文件，需要清除临时文件
     *                 默认有 100 条记录被保存在内存中，如果超过这数量，则最前面的数据被写入临时文件
     *                  如果想自定义内存中的数据量，可以使用 new SXSSFWorkbook (数量)
     *                  官方解释：实现 BigGridDemo 策略的流式 XSSFWorkbook 版本，允许写入非常大的文件而不会耗尽内存，因为任何时候只有可配置的行被部分保存在内存中。
     *                   但仍有可能消耗大量内存，这些内容基于正在使用的操作，如 合并区域，注释.... 仍在只存储在内存中。
     *                    因此，若广泛使用，可能需要大量内存
     * @Date 15:50 2020/7/6
     * @Param []
     * @return void
     **/
    @Test
    public void testWriteVersion07BigDataSuper() throws IOException{

        Long begin = System.currentTimeMillis();

        // 使用 加速 改进的类
        Workbook workbook = new SXSSFWorkbook();

        Sheet sheet = workbook.createSheet("插入大批量数据");

        for (int rowNum = 0; rowNum < 100000; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int colNum = 0; colNum < 10; colNum++) {
                Cell cell = row.createCell(colNum);
                cell.setCellValue(colNum + 1);
            }
        }

        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "/测试07版本的大量数据Excel.xlsx");

        workbook.write(fileOutputStream);

        fileOutputStream.close();

        // 清除临时文件
        ((SXSSFWorkbook) workbook).dispose();

        Long end = System.currentTimeMillis();

        System.out.println("耗时 ---> " + (end - begin) / 1000.0);
    }


}
