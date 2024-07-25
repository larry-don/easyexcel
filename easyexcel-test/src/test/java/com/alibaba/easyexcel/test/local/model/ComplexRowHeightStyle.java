package com.alibaba.easyexcel.test.local.model;

import lombok.Data;

/**
 * @Description: 本类用于
 * @Author larry
 * @Date 2023/7/30 22:43
 */
@Data
public class ComplexRowHeightStyle {
    private Short headRowHeight;
    private Short contentRowHeight;
    private int relativeRowIndex;

    public ComplexRowHeightStyle() {
    }

    public ComplexRowHeightStyle(Short headRowHeight, Short contentRowHeight, int relativeRowIndex) {
        this.headRowHeight = headRowHeight;
        this.contentRowHeight = contentRowHeight;
        this.relativeRowIndex = relativeRowIndex;
    }
}
