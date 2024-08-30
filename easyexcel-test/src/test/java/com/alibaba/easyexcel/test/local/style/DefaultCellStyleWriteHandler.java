package com.alibaba.easyexcel.test.local.style;

import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.handler.context.CellWriteHandlerContext;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import lombok.RequiredArgsConstructor;

/**
 * @Description: 本类用于
 * @Author larry
 * @Date 2024/8/7 17:42
 */

@RequiredArgsConstructor
public class DefaultCellStyleWriteHandler implements CellWriteHandler {

    private final WriteCellStyle contentStyle;

    @Override
    public void afterCellDispose(CellWriteHandlerContext context) {
        WriteCellData<?> firstCellData = context.getFirstCellData();
        firstCellData.setWriteCellStyle(contentStyle);
        //WriteCellStyle.merge(contentStyle, firstCellData.getOrCreateStyle());
    }
}
