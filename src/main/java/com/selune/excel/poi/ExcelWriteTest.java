package com.selune.excel.poi;

import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

/**
 * @author xiaoyunpeng
 * @date 2021/5/12
 */
public class ExcelWriteTest {

    private final static String PATH = "./src/main/resources/excel/";

    @Test
    public void testWriteExcel03() throws IOException {
        // 1. 创建一个工作簿对象03
        Workbook workbook = new HSSFWorkbook();
        // 2. 创建一个工作表
        Sheet sheet1 = workbook.createSheet("03测试1");
        // 3. 创建一个行
        Row row1 = sheet1.createRow(0);
        Row row2 = sheet1.createRow(1);
        // 4. 创建一个单元格
        // (1, 1)
        Cell row1Cell1 = row1.createCell(0);
        row1Cell1.setCellValue("今日新增");
        // (1, 2)
        Cell row1Cell2 = row1.createCell(1);
        row1Cell2.setCellValue(666);
        Cell row2Cell1 = row2.createCell(0);
        row2Cell1.setCellValue("统计时间");
        Cell row2Cell2 = row2.createCell(1);
        row2Cell2.setCellValue(new DateTime().toString("yyyy-MM-dd HH:mm:ss"));

        // 生成一张表 文件流
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "testWriteExcel03.xls");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }

    @Test
    public void testWriteExcel07() throws IOException {
        // 1. 创建一个工作簿对象07
        Workbook workbook = new XSSFWorkbook();
        // 2. 创建一个工作表
        Sheet sheet1 = workbook.createSheet("07测试1");
        // 3. 创建一个行
        Row row1 = sheet1.createRow(0);
        Row row2 = sheet1.createRow(1);
        // 4. 创建一个单元格
        // (1, 1)
        Cell row1Cell1 = row1.createCell(0);
        row1Cell1.setCellValue("今日新增");
        // (1, 2)
        Cell row1Cell2 = row1.createCell(1);
        row1Cell2.setCellValue(666);
        Cell row2Cell1 = row2.createCell(0);
        row2Cell1.setCellValue("统计时间");
        Cell row2Cell2 = row2.createCell(1);
        row2Cell2.setCellValue(new DateTime().toString("yyyy-MM-dd HH:mm:ss"));

        // 生成一张表 文件流
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "testWriteExcel07.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
    }

    @Test
    public void testWriteExcel03BigData() throws IOException {
        Long begin = System.currentTimeMillis();
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        for (int rowNum = 0; rowNum < 65536; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        Long end1 = System.currentTimeMillis();
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "testWriteExcel03BigData.xls");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        Long end2 = System.currentTimeMillis();
        System.out.println((double) (end1 - begin) / 1000);
        System.out.println((double) (end2 - begin) / 1000);
    }

    @Test
    public void testWriteExcel07BigData() throws IOException {
        Long begin = System.currentTimeMillis();
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        for (int rowNum = 0; rowNum < 65537; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        Long end1 = System.currentTimeMillis();
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "testWriteExcel07BigData.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        Long end2 = System.currentTimeMillis();
        System.out.println((double) (end1 - begin) / 1000);
        System.out.println((double) (end2 - begin) / 1000);
    }

    @Test
    public void testWriteExcel07BigDataS() throws IOException {
        Long begin = System.currentTimeMillis();
        Workbook workbook = new SXSSFWorkbook(10000);
        Sheet sheet = workbook.createSheet();
        for (int rowNum = 0; rowNum < 100000; rowNum++) {
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        Long end1 = System.currentTimeMillis();
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "testWriteExcel07BigDataS.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        // 删除临时文件
        ((SXSSFWorkbook) workbook).dispose()
        Long end2 = System.currentTimeMillis();
        System.out.println((double) (end1 - begin) / 1000);
        System.out.println((double) (end2 - begin) / 1000);
    }

}
