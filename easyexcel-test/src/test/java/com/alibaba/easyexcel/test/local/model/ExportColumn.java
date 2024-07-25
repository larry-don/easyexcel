package com.alibaba.easyexcel.test.local.model;

import lombok.Data;
import lombok.RequiredArgsConstructor;

/**
 * @Description: 本类用于表头列
 * @Author larry
 * @Date 2023/7/30 21:39
 */
@Data
public class ExportColumn {
    private String columnName;
    private String columnCode;
    private int index;

    public ExportColumn(String columnName, String columnCode, int index) {
        this.columnName = columnName;
        this.columnCode = columnCode;
        this.index = index;
    }
}
