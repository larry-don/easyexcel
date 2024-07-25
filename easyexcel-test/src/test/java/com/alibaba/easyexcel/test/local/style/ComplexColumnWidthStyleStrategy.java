package com.alibaba.easyexcel.test.local.style;

import com.alibaba.easyexcel.test.local.model.ComplexColumnWidthStyle;
import com.alibaba.easyexcel.test.local.model.ComplexRowHeightStyle;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.handler.context.CellWriteHandlerContext;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.style.column.AbstractColumnWidthStyleStrategy;
import lombok.Data;
import org.apache.poi.ss.usermodel.Cell;

import java.util.List;
import java.util.concurrent.ArrayBlockingQueue;

/**
 * @Description: 本类用于
 * @Author larry
 * @Date 2023/8/9 22:49
 */

@Data
public class ComplexColumnWidthStyleStrategy extends AbstractColumnWidthStyleStrategy {

    private ArrayBlockingQueue<ComplexColumnWidthStyle> complexColumnWidthStyles;

    @Override
    protected void setColumnWidth(WriteSheetHolder writeSheetHolder, List<WriteCellData<?>> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        if (!complexColumnWidthStyles.isEmpty()){
            ComplexColumnWidthStyle peek = complexColumnWidthStyles.peek();
            int columnIndex = cell.getColumnIndex();
            if (columnIndex== peek.getColumnIndex()){
               writeSheetHolder.getSheet().setColumnWidth(columnIndex,peek.getWidth());
               complexColumnWidthStyles.poll();
            }
        }
    }
}
