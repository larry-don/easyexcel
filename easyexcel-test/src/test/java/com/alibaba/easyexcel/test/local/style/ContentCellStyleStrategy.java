package com.alibaba.easyexcel.test.local.style;

import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.handler.context.CellWriteHandlerContext;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.Objects;

/**
 * @Description: 本类用于
 * @Author larry
 * @Date 2024/8/7 14:38
 */
public class ContentCellStyleStrategy implements CellWriteHandler {

    @Override
    public void afterCellDispose(CellWriteHandlerContext context) {
        setContentCellStyle(context);
    }

    protected void setContentCellStyle(CellWriteHandlerContext context) {
        WriteCellStyle writeCellStyle = new WriteCellStyle();
        if (Objects.equals(1, context.getColumnIndex()) && Objects.equals(0, context.getRelativeRowIndex())) {
            writeCellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            WriteFont writeFont = new WriteFont();
            writeFont.setColor(IndexedColors.RED.getIndex());
            writeFont.setBold(true);
            writeFont.setFontHeightInPoints((short) 15);
            writeCellStyle.setWriteFont(writeFont);
            WriteCellData<?> firstCellData = context.getFirstCellData();
            //当存在多个自定义CellWriteHandler时，需要用这种方式设置style；使用WriteCellStyle.merge会存在writeStyle被污染的问题，导致会被最后一个自定义格式覆盖
            firstCellData.setWriteCellStyle(writeCellStyle);
            //WriteCellStyle.merge(writeCellStyle, firstCellData.getOrCreateStyle());
        }
    }
}
