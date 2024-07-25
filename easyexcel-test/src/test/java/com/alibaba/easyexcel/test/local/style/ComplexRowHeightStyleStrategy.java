package com.alibaba.easyexcel.test.local.style;

import com.alibaba.easyexcel.test.local.model.ComplexRowHeightStyle;
import com.alibaba.excel.write.style.row.AbstractRowHeightStyleStrategy;
import lombok.Data;
import org.apache.poi.ss.usermodel.Row;

import java.util.concurrent.ArrayBlockingQueue;

/**
 * @Description: 本类用于
 * @Author larry
 * @Date 2023/7/30 22:59
 */
@Data
public class ComplexRowHeightStyleStrategy extends AbstractRowHeightStyleStrategy {

    private ArrayBlockingQueue<ComplexRowHeightStyle> complexRowHeightStyles;

    @Override
    protected void setHeadColumnHeight(Row row, int relativeRowIndex) {
        ComplexRowHeightStyle complexRowHeightStyle = complexRowHeightStyles.peek();
        if (complexRowHeightStyle.getRelativeRowIndex() == relativeRowIndex) {
            row.setHeightInPoints(complexRowHeightStyle.getHeadRowHeight());
            complexRowHeightStyles.poll();
        }
    }

    @Override
    protected void setContentColumnHeight(Row row, int relativeRowIndex) {
        if (complexRowHeightStyles!=null&&!complexRowHeightStyles.isEmpty()){
            ComplexRowHeightStyle complexRowHeightStyle = complexRowHeightStyles.peek();
            if (complexRowHeightStyle.getRelativeRowIndex() == relativeRowIndex) {
                row.setHeightInPoints(complexRowHeightStyle.getContentRowHeight());
                complexRowHeightStyles.poll();
            }
        }
    }
}
