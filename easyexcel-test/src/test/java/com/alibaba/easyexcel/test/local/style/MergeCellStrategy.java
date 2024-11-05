package com.alibaba.easyexcel.test.local.style;

import com.alibaba.excel.write.handler.RowWriteHandler;
import com.alibaba.excel.write.handler.context.RowWriteHandlerContext;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

/**
 * @Description: 本类用于
 * @Author larry
 * @Date 2024/10/12 10:59
 */
public class MergeCellStrategy implements RowWriteHandler {

    @Override
    public void afterRowDispose(RowWriteHandlerContext context) {
        Sheet sheet = context.getWriteSheetHolder().getSheet();
        if (!isCellRangeMerged(sheet,0,0,0,2)) {
            CellRangeAddress cellAddresses = new CellRangeAddress(0, 0, 0, 2);
            sheet.addMergedRegion(cellAddresses);
        }
    }


    private boolean isCellRangeMerged(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        CellRangeAddress targetRange = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);

        return mergedRegions.stream()
            .anyMatch(region ->
                region.isInRange(targetRange.getFirstRow(), targetRange.getFirstColumn()) ||
                    region.isInRange(targetRange.getLastRow(), targetRange.getLastColumn()));
    }


}
