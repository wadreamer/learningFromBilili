package poi;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Date;

/**
 * ClassName: ExcelWriteTest
 * Description:
 * date: 2020/7/6 14:47
 *
 * @author CFG
 * @since JDK 1.8
 */
public class ExcelReadTest {

    private final static String PATH = "D:/IntelliJ IDEA 2019.3.3/idea_workplace/excel_poi/src/main/resources";

    /*
     * @Author wadreamer
     * @Description //TODO 注意获取值的类型，否则会报错
     * @Date 16:29 2020/7/6
     * @Param []
     * @return void
     **/
    @Test
    public void testReadVersion03() throws IOException {

        FileInputStream fileInputStream = new FileInputStream(PATH + "/测试03版本的Excel.xls");

        Workbook workbook = new HSSFWorkbook(fileInputStream);

        Sheet sheet = workbook.getSheetAt(0);

        Row row1 = sheet.getRow(0);
        Cell cell11 = row1.getCell(0);
        System.out.println(cell11.getStringCellValue());

        Cell cell12 = row1.getCell(1);
        System.out.println(cell12.getNumericCellValue());

        fileInputStream.close();
    }

    @Test
    public void testReadVersion07() throws IOException {

        FileInputStream fileInputStream = new FileInputStream(PATH + "/测试07版本的Excel.xlsx");

        Workbook workbook = new XSSFWorkbook(fileInputStream);

        Sheet sheet = workbook.getSheetAt(0);

        Row row1 = sheet.getRow(0);
        Cell cell11 = row1.getCell(0);
        System.out.println(cell11.getStringCellValue());

        Cell cell12 = row1.getCell(1);
        System.out.println(cell12.getNumericCellValue());

        fileInputStream.close();
    }

    /*
     * @Author wadreamer
     * @Description //TODO 正确处理单元格内不同的数据类型
     * @Date 16:50 2020/7/6
     * @Param []
     * @return void
     **/
    @Test
    public void testCellType() throws IOException {
        FileInputStream fileInputStream = new FileInputStream(PATH + "/滞纳金统计分析.xls");

        Workbook workbook = new HSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);

        // 获取标题行
        Row titleRow = sheet.getRow(1);
        if (titleRow != null) {
            // 获取该行的单元格数量
            int cellCount = titleRow.getPhysicalNumberOfCells();
            for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                Cell cell = titleRow.getCell(cellNum);
                if (cell != null) {
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellValue + " | ");
                }
            }
            System.out.println();
        }

        // 获取表中的数据
        int rowCount = sheet.getPhysicalNumberOfRows();

        for (int rowNum = 2; rowNum < rowCount; rowNum++) {
            Row rowData = sheet.getRow(rowNum);

            if (rowData != null) {
                int cellCount = rowData.getPhysicalNumberOfCells();

                for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                    System.out.print("【" + (rowNum + 1) + "-" + (cellNum + 1) + "】");

                    Cell cell = rowData.getCell(cellNum);
                    if(cell != null){
                        int cellType = cell.getCellType();
                        String cellValue ="";

                        switch (cellType){
                            case HSSFCell.CELL_TYPE_STRING: // 字符串
                                System.out.print("【STRING】");
                                cellValue = cell.getStringCellValue();
                                break;
                            case HSSFCell.CELL_TYPE_BOOLEAN: // 布尔
                                System.out.print("【BOOLEAN】");
                                cellValue = String.valueOf(cell.getBooleanCellValue());
                                break;
                            case HSSFCell.CELL_TYPE_BLANK: // 空
                                System.out.print("【BLANK:】");
                                break;
                            case HSSFCell.CELL_TYPE_NUMERIC: // 数字 (日期，普通数字)
                                System.out.print("【NUMERIC】");
                                if(HSSFDateUtil.isCellDateFormatted(cell)){ // 日期
                                    System.out.print("【日期】");
                                    Date date = cell.getDateCellValue();
                                    cellValue = new DateTime(date).toString("yyyy-MM-dd HH:mm:ss");
                                }else{
                                    System.out.print("转化为字符串输出");
                                    cell.setCellType(HSSFCell.CELL_TYPE_STRING); // 防止数字过长，转化为字符串输出
                                    cellValue = cell.toString();
                                }
                                break;
                            case HSSFCell.CELL_TYPE_ERROR:
                                System.out.print("【数据类型错误】");
                                break;
                        }
                        System.out.println(cellValue);
                    }
                }
            }
        }
        fileInputStream.close();
    }

}
