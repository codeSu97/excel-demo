package com.selune.excel.poi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

/**
 * @author xiaoyunpeng
 * @date 2021/5/12
 */
public class ExcelReadTest {

    private final static String PATH = "./src/main/resources/excel/";

    @Test
    public void testRead03() throws IOException {
        // 0. 获取文件流
        FileInputStream fileInputStream = new FileInputStream(PATH + "testWriteExcel03.xls");
        // 1. 创建一个工作簿
        Workbook workbook = new HSSFWorkbook(fileInputStream);
        // 2. 创建一个工作表
        Sheet sheet = workbook.getSheetAt(0);
        // 3. 创建一个行
        Row row = sheet.getRow(0);
        // 4. 创建一个单元格
        Cell cell = row.getCell(1);
        CellStyle cellStyle = cell.getCellStyle();
        System.out.println(cellStyle.getFillForegroundColor());
        System.out.println(cellStyle.getFillBackgroundColor());
        System.out.println(cell.getNumericCellValue());
        fileInputStream.close();
    }

    @Test
    public void testFormula() throws IOException {
        FileInputStream fileInputStream = new FileInputStream(PATH + "公式.xlsx");
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(4);
        Cell cell = row.getCell(0);

        // 计算公式
        FormulaEvaluator formulaEvaluator = new XSSFFormulaEvaluator((XSSFWorkbook) workbook);

        // 输出单元格内容
        CellType cellType = cell.getCellType();
        switch (cellType) {
            case FORMULA:
                String formula = cell.getCellFormula();
                System.out.println(formula);

                // 计算
                CellValue evaluate = formulaEvaluator.evaluate(cell);
                String cellValue = evaluate.formatAsString();
                System.out.println(cellValue);
                break;
            case NUMERIC:
                double numericCellValue = cell.getNumericCellValue();
                System.out.println(numericCellValue);
                break;
        }
    }
}
