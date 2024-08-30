package com.alibaba.easyexcel.test.local.style;

import com.alibaba.easyexcel.test.local.model.ComplexCellStyle;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.handler.context.CellWriteHandlerContext;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.AbstractCellStyleStrategy;
import lombok.Data;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.Objects;
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
        /*WriteCellStyle writeCellStyle = new WriteCellStyle();
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
        }*/
    }

    @Override
    protected void setContentCellStyle(CellWriteHandlerContext context) {
        if (cellStyles != null && !cellStyles.isEmpty()) {
            ComplexCellStyle style = cellStyles.peek();
            if (Objects.equals(style.getColumnIndex(), context.getColumnIndex()) && Objects.equals(style.getRowIndex(), context.getRelativeRowIndex())) {
                WriteCellStyle writeCellStyle = buildDefaultStyle();
                if (style.getFillForegroundColor() != null) {
                    writeCellStyle.setFillForegroundColor(style.getFillForegroundColor());
                }
                if (style.getFillPatternType() != null) {
                    writeCellStyle.setFillPatternType(style.getFillPatternType());
                }
                WriteFont writeFont = writeCellStyle.getWriteFont();
                if (style.getFontColor() != null) {
                    writeFont.setColor(style.getFontColor());
                }
                if (style.getBold() != null) {
                    writeFont.setBold(style.getBold());
                }
                WriteCellData<?> firstCellData = context.getFirstCellData();
                firstCellData.setWriteCellStyle(writeCellStyle);
                cellStyles.poll();
            }
        }
    }

    public WriteCellStyle buildDefaultStyle() {
        WriteCellStyle writeCellStyle = new WriteCellStyle();
        writeCellStyle.setWrapped(true);
        writeCellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        writeCellStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
        writeCellStyle.setBorderTop(BorderStyle.THIN);
        writeCellStyle.setBorderBottom(BorderStyle.THIN);
        writeCellStyle.setBorderLeft(BorderStyle.THIN);
        writeCellStyle.setBorderRight(BorderStyle.THIN);
        writeCellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        writeCellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        writeCellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        writeCellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        writeCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        WriteFont writeFont = new WriteFont();
        writeFont.setFontName("SimSun");
        writeFont.setFontHeightInPoints((short) 10);
        writeCellStyle.setWriteFont(writeFont);
        return writeCellStyle;
    }
}
