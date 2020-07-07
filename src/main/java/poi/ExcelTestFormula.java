package poi;

import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.IOException;

/**
 * ClassName: ExcelTestFormula
 * Description:
 * date: 2020/7/6 18:01
 *
 * @author CFG
 * @since JDK 1.8
 */
public class ExcelTestFormula {

    private final static String PATH = "D:/IntelliJ IDEA 2019.3.3/idea_workplace/excel_poi/src/main/resources";

    @Test
    public void testFormula() throws IOException {
        FileInputStream fileInputStream = new FileInputStream(PATH + "/公式.xls");

        // 面向接口编程，修改的时候，只需要修改接口
        Workbook workbook = new HSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);

        Row row = sheet.getRow(4);
        Cell formulaCell = row.getCell(0);

        // 得到计算器 eval
        FormulaEvaluator formulaEvaluator = new HSSFFormulaEvaluator((HSSFWorkbook) workbook);

        // 输出单元格内容
        int cellType = formulaCell.getCellType();

        switch (cellType){
            case Cell.CELL_TYPE_FORMULA:
                String formula = formulaCell.getCellFormula(); // 计算公式
                System.out.println(formula);

                CellValue evaluate = formulaEvaluator.evaluate(formulaCell); // 利用 公式计算器 计算 公式结果
                String cellValue = evaluate.formatAsString();
                System.out.println(cellValue);
                break;
        }
        fileInputStream.close();
    }
}
