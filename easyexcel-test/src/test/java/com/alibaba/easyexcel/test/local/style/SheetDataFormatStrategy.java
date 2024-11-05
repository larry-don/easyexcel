package com.alibaba.easyexcel.test.local.style;

import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.stream.IntStream;

/**
 * @Description: 本类用于
 * @Author larry
 * @Date 2024/11/4 17:27
 */
public class SheetDataFormatStrategy implements SheetWriteHandler {
    private CellStyle cellStyle;

    @Override
    public void beforeSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {

    }

    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        Sheet sheet = writeSheetHolder.getSheet();
        Workbook workbook = writeWorkbookHolder.getWorkbook();
        if (cellStyle == null) {
            cellStyle = workbook.createCellStyle();
            DataFormat dataFormat = workbook.createDataFormat();
            cellStyle.setDataFormat(dataFormat.getFormat("@"));
        }
        IntStream.range(0, 3).forEach(i -> sheet.setDefaultColumnStyle(i, cellStyle));
    }
}
