package com.alibaba.easyexcel.test.demo.write;

import com.alibaba.easyexcel.test.local.style.CellStyleWriteHandle;
import com.alibaba.easyexcel.test.local.style.MergeCellStrategy;
import com.alibaba.easyexcel.test.local.style.SheetDataFormatStrategy;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.converters.date.DateStringConverter;
import com.alibaba.excel.util.ListUtils;
import com.alibaba.excel.write.metadata.WriteSheet;
import org.junit.jupiter.api.Test;

import java.math.BigDecimal;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Random;

/**
 * @Description: 本类用于
 * @Author larry
 * @Date 2024/10/12 10:37
 */
public class WriteTest1 {
    @Test
    public void testDynamicHead() {
        String fileName = "C:/Users/larry/Desktop/模版测试/" + "_导入模板_" + System.currentTimeMillis() + ".xlsx";
        try (ExcelWriter excelWriter = EasyExcel.write(fileName).build()) {
            for (int i = 0; i < 2; i++) {
                WriteSheet writeSheet = EasyExcel.writerSheet(i, "模板" + i).automaticMergeHead(false).
                    registerConverter(new DateStringConverter()).build();
                if (i == 0) {
                    writeSheet.setHead(head());
                    writeSheet.setCustomWriteHandlerList(Arrays.asList(new CellStyleWriteHandle(), new MergeCellStrategy()));
                    excelWriter.write(dataList(), writeSheet);
                }
                if (i == 1) {
                    writeSheet.setHead(head1());
                    writeSheet.setCustomWriteHandlerList(Arrays.asList(new CellStyleWriteHandle()));
                    excelWriter.write(dataList1(), writeSheet);
                }
            }
        }
    }

    @Test
    public void testSheetDateFormat() {
        String fileName = "C:/Users/larry/Desktop/模版测试/" + "_导入模板_" + System.currentTimeMillis() + ".xlsx";
        try (ExcelWriter excelWriter = EasyExcel.write(fileName).build()) {
            for (int i = 0; i < 2; i++) {
                WriteSheet writeSheet = EasyExcel.writerSheet(i, "模板" + i)
                    .automaticMergeHead(false)
                    //.registerConverter(new StringStringConverter())
                    //.registerConverter(new BigDecimalStringConverter())
                    //.registerConverter(new LongStringConverter())
                    .build();
                if (i == 0) {
                    writeSheet.setHead(head());
                    writeSheet.setCustomWriteHandlerList(Arrays.asList(new SheetDataFormatStrategy(),new CellStyleWriteHandle()));
                    excelWriter.write(dataList(), writeSheet);
                }
                if (i == 1) {
                    writeSheet.setHead(head1());
                    //writeSheet.setCustomWriteHandlerList(Arrays.asList(new CellStyleWriteHandle()));
                    excelWriter.write(dataList1(), writeSheet);
                }
            }
        }
    }

    private List<List<String>> head() {
        List<List<String>> list = ListUtils.newArrayList();
        List<String> head0 = ListUtils.newArrayList();
        head0.add("字符串test");
        head0.add("cuikai");
        List<String> head1 = ListUtils.newArrayList();
        head1.add("数字test" + System.currentTimeMillis());
        head1.add("cuikai");
        List<String> head2 = ListUtils.newArrayList();
        head2.add("日期test" + System.currentTimeMillis());
        head2.add("cuikai");
        list.add(head0);
        list.add(head1);
        list.add(head2);
        return list;
    }

    private List<List<Object>> dataList() {
        List<List<Object>> list = ListUtils.newArrayList();
        for (int i = 0; i < 10; i++) {
            if (i == 0) {
                List<Object> data0 = ListUtils.newArrayList();
                data0.add("project");
                data0.add("price");
                data0.add("date");
                list.add(data0);
                List<Object> data1 = ListUtils.newArrayList();
                data1.add("项目");
                data1.add("金额");
                data1.add("日期");
                list.add(data1);
                continue;
            }
            List<Object> data = ListUtils.newArrayList();
            data.add(true);
            BigDecimal bigDecimal = new BigDecimal(1245120.0212157487);
            BigDecimal bigDecimal1 = bigDecimal.setScale(8, BigDecimal.ROUND_HALF_UP);
            data.add(bigDecimal1);
            data.add("2024-11-05 09:36:29");
            list.add(data);
        }
        return list;
    }

    private List<List<Object>> dataList1() {
        List<List<Object>> list = ListUtils.newArrayList();
        for (int i = 0; i < 10; i++) {
            List<Object> data = ListUtils.newArrayList();
            data.add("姓名" + i);
            data.add(new Random(100).nextDouble());
            data.add(new Date());
            list.add(data);
        }
        return list;
    }

    private List<List<String>> head1() {
        List<List<String>> list = ListUtils.newArrayList();
        List<String> head0 = ListUtils.newArrayList();
        head0.add("姓名test" + System.currentTimeMillis());
        head0.add("ck测试姓名");
        List<String> head1 = ListUtils.newArrayList();
        head1.add("年龄test" + System.currentTimeMillis());
        head1.add("ck测试年龄");
        List<String> head2 = ListUtils.newArrayList();
        head2.add("出生日期test" + System.currentTimeMillis());
        head2.add("ck出生日期");
        list.add(head0);
        list.add(head1);
        list.add(head2);
        return list;
    }
}
