package com.alibaba.easyexcel.test.local.model;

import lombok.Data;

import java.util.List;

/**
 * @Description: 本类用于
 * @Author larry
 * @Date 2023/7/30 21:44
 */

@Data
public class ExportWorkBook {
    private String fileName;
    private List<ExportSheet> sheets;
}
