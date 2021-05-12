package com.selune.excel.easy;

import com.alibaba.excel.EasyExcel;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import org.junit.Test;

public class TestExcelWrite {

    private final static String PATH = "./src/main/resources/excel/";

    // 根据list写入Excel
    @Test
    public void simpleTest() {
        String fileName = PATH + "EasyExcel.xlsx";
        EasyExcel.write(fileName, DemoData.class).sheet().doWrite(data());

    }

    private List<DemoData> data() {
        List<DemoData> list = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            DemoData demoData = new DemoData();
            demoData.setString("字符串" + i);
            demoData.setDate(new Date());
            demoData.setADouble(1.33);
            demoData.setIgnore("Ignore" + i);
            list.add(demoData);
        }
        return list;
    }
}
