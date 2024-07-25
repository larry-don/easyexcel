package com.alibaba.easyexcel.test.local.style;

import com.alibaba.easyexcel.test.local.model.ComplexCellStyle;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.handler.context.CellWriteHandlerContext;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.AbstractCellStyleStrategy;
import lombok.Data;
import org.apache.poi.ss.usermodel.Cell;

import java.util.List;
import java.util.concurrent.ArrayBlockingQueue;

/**
 * @Description: 本类用于
 * @Author larry
 * @Date 2023/7/30 23:39
 */
@Data
public class ComplexCellStyleStrategy extends AbstractCellStyleStrategy {
    private ArrayBlockingQueue<ComplexCellStyle> cellStyles;

    @Override
    protected void setHeadCellStyle(CellWriteHandlerContext context) {
        WriteCellStyle writeCellStyle = new WriteCellStyle();
        if (cellStyles != null && !cellStyles.isEmpty()) {
            ComplexCellStyle style = cellStyles.peek();
            List<WriteCellData<?>> cellDataList = context.getCellDataList();
            if (style.getColumnIndex() == context.getColumnIndex() && style.getRowIndex() == context.getRelativeRowIndex()) {
                writeCellStyle.setFillForegroundColor(style.getBackgroundColor());
                WriteFont writeFont = new WriteFont();
                writeFont.setColor(style.getColor());
                writeFont.setFontName(style.getFontName());
                writeFont.setFontHeightInPoints(style.getFontHeightInPoints());
                writeFont.setBold(style.getBold());
                writeCellStyle.setWriteFont(writeFont);
                writeCellStyle.setWrapped(style.getWrapped());
                cellStyles.poll();
                WriteCellStyle.merge(writeCellStyle, context.getFirstCellData().getOrCreateStyle());
            }
        }
    }

    @Override
    protected void setContentCellStyle(CellWriteHandlerContext context) {

    }
}
