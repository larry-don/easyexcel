package com.alibaba.easyexcel.test.local.model;

import lombok.Data;

import java.util.List;

/**
 * @Description: 本类用于导出excel sheet页
 * @Author larry
 * @Date 2023/7/30 21:42
 */
@Data
public class ExportSheet {
    private String sheetName;
    private String sheetCode;
    private int num;
    private List<ExportColumn> columns;

    public ExportSheet(String sheetName, String sheetCode, int num) {
        this.sheetName = sheetName;
        this.sheetCode = sheetCode;
        this.num = num;
    }
}
