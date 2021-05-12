package com.selune.excel.easy;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.NumberFormat;
import java.util.Date;
import lombok.Data;

@Data
public class DemoData {
    @ExcelProperty("字符串标题")
    private String string;

    @ExcelProperty("日期标题")
    private Date date;

    @ExcelProperty("数字标题")
    @NumberFormat("##.00")
    private Double aDouble;

    /** 忽略这个字段 */
    @ExcelIgnore
    private String ignore;
}
