package com.alibaba.easyexcel.test.local.style;

import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.handler.context.CellWriteHandlerContext;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import org.apache.commons.collections4.CollectionUtils;

import java.util.List;

/**
 * @Description: 本类用于
 * @Author larry
 * @Date 2024/8/29 13:39
 */
public class DefaultCellStyleStrategy extends HorizontalCellStyleStrategy {


    @Override
    protected void setHeadCellStyle(CellWriteHandlerContext context) {
    }

    @Override
    protected void setContentCellStyle(CellWriteHandlerContext context) {
        List<WriteCellStyle> contentWriteCellStyleList = getContentWriteCellStyleList();
        if (stopProcessing(context) || CollectionUtils.isEmpty(contentWriteCellStyleList)) {
            return;
        }
        WriteCellData<?> cellData = context.getFirstCellData();
        int size = 2;
        WriteCellStyle currentStyle = context.getRelativeRowIndex() < size ? contentWriteCellStyleList.get(0) : contentWriteCellStyleList.get(1);
        cellData.setWriteCellStyle(currentStyle);
    }
}
