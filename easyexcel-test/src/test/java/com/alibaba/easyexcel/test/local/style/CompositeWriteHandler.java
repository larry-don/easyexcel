package com.alibaba.easyexcel.test.local.style;

import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.handler.context.CellWriteHandlerContext;
import lombok.RequiredArgsConstructor;

import java.util.List;

/**
 * @Description: 本类用于
 * @Author larry
 * @Date 2024/8/7 15:13
 */
@RequiredArgsConstructor
public class CompositeWriteHandler implements CellWriteHandler {

    private final List<CellWriteHandler> CellWriteHandlers;

    @Override
    public void afterCellDispose(CellWriteHandlerContext context) {
        for (CellWriteHandler cellWriteHandler : CellWriteHandlers) {
            cellWriteHandler.afterCellDispose(context);
        }
    }
}
