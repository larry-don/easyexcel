package com.alibaba.easyexcel.test.local.style;

import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.handler.context.CellWriteHandlerContext;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @Description: 本类用于
 * @Author larry
 * @Date 2024/7/25 15:37
 */
public class CellStyleWriteHandle implements CellWriteHandler {

    private CellStyle cellStyle;

    /*@Override
    public void afterCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        //设置单元格格式为文本
        Workbook workbook = writeSheetHolder.getSheet().getWorkbook();
        if (cellStyle == null) {
            cellStyle = workbook.createCellStyle();
            DataFormat dataFormat = workbook.createDataFormat();
            cellStyle.setDataFormat(dataFormat.getFormat("@"));
        }
        cell.setCellStyle(cellStyle);
    }*/

    @Override
    public void afterCellDispose(CellWriteHandlerContext context) {
        Boolean head = context.getHead();
        Cell cell = context.getCell();
        //设置单元格格式为文本
        Workbook workbook = context.getWriteWorkbookHolder().getWorkbook();
        if (cellStyle == null) {
            cellStyle = workbook.createCellStyle();
            DataFormat dataFormat = workbook.createDataFormat();
            cellStyle.setDataFormat(dataFormat.getFormat("@"));
        }
        cell.setCellStyle(cellStyle);
    }

    /*@Override
    public void afterCellCreate(CellWriteHandlerContext context) {
        Boolean head = context.getHead();
        Cell cell = context.getCell();
        //设置单元格格式为文本
        Workbook workbook = context.getWriteWorkbookHolder().getWorkbook();
        if (cellStyle == null) {
            cellStyle = workbook.createCellStyle();
            DataFormat dataFormat = workbook.createDataFormat();
            cellStyle.setDataFormat(dataFormat.getFormat("@"));
        }
        cell.setCellStyle(cellStyle);
    }*/
}
