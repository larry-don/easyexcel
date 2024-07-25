package com.alibaba.easyexcel.test.local.style;

import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @Description: 本类用于
 * @Author larry
 * @Date 2023/7/31 14:14
 */
public class HiddenSheetWriteHandler  implements SheetWriteHandler {
    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        Workbook workbook = writeWorkbookHolder.getWorkbook();
        Integer sheetNo = writeSheetHolder.getSheetNo();
        workbook.setSheetHidden(sheetNo,true);
    }
}
