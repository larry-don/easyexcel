package com.alibaba.easyexcel.test.local.style;

import com.alibaba.excel.converters.WriteConverterContext;
import com.alibaba.excel.converters.string.StringStringConverter;
import com.alibaba.excel.metadata.data.WriteCellData;

/**
 * @Description: 本类用于
 * @Author larry
 * @Date 2024/11/4 17:02
 */
public class StringConverter extends StringStringConverter {

    @Override
    public WriteCellData<?> convertToExcelData(WriteConverterContext<String> context) throws Exception {
        return super.convertToExcelData(context);
    }
}
