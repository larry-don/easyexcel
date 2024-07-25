package com.alibaba.easyexcel.test.local.model;

import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * @Description: 本类用于
 * @Author larry
 * @Date 2023/8/10 14:13
 */

@Data
@NoArgsConstructor
public class ComplexColumnWidthStyle {
    private int width;
    private int columnIndex;

    public ComplexColumnWidthStyle(int width, int columnIndex) {
        this.width = width;
        this.columnIndex = columnIndex;
    }
}
