package com.selune.excel.easy;

import com.alibaba.excel.EasyExcel;
import org.junit.Test;

public class TestExcelRead {

    private final static String PATH = "./src/main/resources/excel/";

    @Test
    public void read() {
        String filePath = PATH + "EasyExcel.xlsx";
        EasyExcel.read(filePath).sheet(0).doRead();
        EasyExcel.read(filePath, DemoData.class, new DemoDataListener()).sheet(0).doRead();
    }
}
