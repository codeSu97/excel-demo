package com.selune.excel.poi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
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

    }
}
