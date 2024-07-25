package com.alibaba.easyexcel.test.local.style;

import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import lombok.AllArgsConstructor;
import lombok.Data;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * @Description: 本类用于
 * @Author larry
 * @Date 2023/7/31 15:48
 */
@Data
@AllArgsConstructor
public class HiddenColumnWriteStrategy implements SheetWriteHandler{

    private int columnIndex;

    @Override
    public void afterSheetCreate(WriteWorkbookHolder writeWorkbookHolder, WriteSheetHolder writeSheetHolder) {
        Sheet sheet = writeSheetHolder.getSheet();
        sheet.setColumnHidden(columnIndex,true);
    }
}
