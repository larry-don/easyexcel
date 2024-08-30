package com.alibaba.easyexcel.test.local.style;

import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @Description: 本类用于
 * @Author larry
 * @Date 2024/8/29 13:48
 */
public class CellFormatStrategy implements CellWriteHandler {

    @Override
    public void afterCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        //设置单元格格式为文本
        Workbook workbook = writeSheetHolder.getSheet().getWorkbook();
        CellStyle cellStyle = workbook.createCellStyle();
        DataFormat dataFormat = workbook.createDataFormat();
        cellStyle.setDataFormat(dataFormat.getFormat("@"));
        cell.setCellStyle(cellStyle);
    }
}
