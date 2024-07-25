package com.alibaba.easyexcel.test.local.model;

import lombok.Data;

/**
 * @Description: 本类用于
 * @Author larry
 * @Date 2023/7/30 23:33
 */
@Data
public class ComplexCellStyle {
    private Integer rowIndex;
    private Integer columnIndex;

    /*背景颜色*/
    private Short backgroundColor;

    /*字体名称*/
    private String fontName;

    /*自动换行*/
    private Boolean wrapped;

    /*是否加粗*/
    private Boolean bold;

    /*字号*/
    private Short fontHeightInPoints;

    /*字体颜色*/
    private Short color;



    public ComplexCellStyle() {
    }

    public ComplexCellStyle(Integer rowIndex, Integer columnIndex, Short backgroundColor) {
        this.rowIndex = rowIndex;
        this.columnIndex = columnIndex;
        this.backgroundColor = backgroundColor;
    }
}
