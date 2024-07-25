package com.alibaba.easyexcel.test.demo.write;

import com.alibaba.easyexcel.test.local.model.*;
import com.alibaba.easyexcel.test.local.style.ComplexCellStyleStrategy;
import com.alibaba.easyexcel.test.local.style.ComplexColumnWidthStyleStrategy;
import com.alibaba.easyexcel.test.local.style.ComplexRowHeightStyleStrategy;
import com.alibaba.easyexcel.test.util.TestFileUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.util.ListUtils;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;
import java.util.concurrent.ArrayBlockingQueue;

/**
 * @Description: 本类用于动态导出表头，表体内容测试类
 * @Author larry
 * @Date 2023/7/30 21:34
 */
public class DynamicWriteTest {

    public ExportWorkBook buildWorkBook() {
        ExportWorkBook workBook = new ExportWorkBook();
        String fileName = TestFileUtil.getPath() + "动态导出样式测试ck" + System.currentTimeMillis() + ".xlsx";
        workBook.setFileName(fileName);
        ExportColumn exportColumn1 = new ExportColumn("姓名", "name", 0);
        ExportColumn exportColumn2 = new ExportColumn("年龄", "age", 1);
        ExportColumn exportColumn3 = new ExportColumn("班级", "class", 2);
        ExportColumn exportColumn4 = new ExportColumn("英语", "English", 0);
        ExportColumn exportColumn5 = new ExportColumn("语文", "Chinese", 1);
        ExportColumn exportColumn6 = new ExportColumn("数学", "math", 2);
        ExportColumn exportColumn7 = new ExportColumn("政治", "politics", 3);
        List<ExportColumn> columns1 = Arrays.asList(exportColumn1, exportColumn2, exportColumn3);
        List<ExportColumn> columns2 = Arrays.asList(exportColumn4, exportColumn5, exportColumn6, exportColumn7);
        ExportSheet exportSheet1 = new ExportSheet("个人信息", "personalInfo", 0);
        ExportSheet exportSheet2 = new ExportSheet("成绩", "score", 1);
        exportSheet1.setColumns(columns1);
        exportSheet2.setColumns(columns2);
        List<ExportSheet> exportSheets = Arrays.asList(exportSheet1, exportSheet2);
        workBook.setSheets(exportSheets);
        return workBook;
    }


    @Test
    /**
     * @return: void
     * @Description: 动态表头生成、格式控制
     * @Author: larry
     * @Date: 2023/7/31 9:34
     **/
    public void testDynamicHead() {
        ExportWorkBook workBook = buildWorkBook();
        String fileName = workBook.getFileName();
        try (ExcelWriter excelWriter = EasyExcel.write(fileName).build()) {
            List<ExportSheet> sheets = workBook.getSheets();
            for (int i = 0; i < sheets.size(); i++) {
                ExportSheet sheet = sheets.get(i);
                WriteSheet writeSheet = EasyExcel.writerSheet(i, sheet.getSheetName()).build();
                List<List<String>> head = getHead(sheet);
                writeSheet.setHead(head);

                //行高设置
                /*ArrayBlockingQueue<ComplexRowHeightStyle> complexRowHeightStyles = buildRowHeightStyle(head);
                ComplexRowHeightStyleStrategy complexRowHeightStyleStrategy = new ComplexRowHeightStyleStrategy();
                complexRowHeightStyleStrategy.setComplexRowHeightStyles(complexRowHeightStyles);*/

                //样式设置（字体颜色、字号大小、加粗、背景颜色、字体、自动换行）
                //默认样式
                HorizontalCellStyleStrategy horizontalCellStyleStrategy = new HorizontalCellStyleStrategy();
                WriteCellStyle headStyle = new WriteCellStyle();
                WriteCellStyle contentStyle = new WriteCellStyle();

                //表头默认样式
                headStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
                headStyle.setWrapped(true);
                headStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
                WriteFont headFont = new WriteFont();
                headFont.setFontName("微软雅黑");
                headFont.setFontHeightInPoints((short) 12);
                headFont.setColor(IndexedColors.BLACK.getIndex());
                headFont.setBold(false);
                headStyle.setWriteFont(headFont);
                horizontalCellStyleStrategy.setHeadWriteCellStyle(headStyle);

                //内容默认样式
                contentStyle.setWrapped(true);
                contentStyle.setHorizontalAlignment(HorizontalAlignment.LEFT);
                WriteFont contentFont = new WriteFont();
                contentFont.setFontName("微软雅黑");
                contentFont.setFontHeightInPoints((short) 10);
                contentFont.setBold(false);
                contentFont.setColor(IndexedColors.BLACK.getIndex());
                contentStyle.setWriteFont(contentFont);

                horizontalCellStyleStrategy.setHeadWriteCellStyle(headStyle);
                horizontalCellStyleStrategy.setContentWriteCellStyleList(Arrays.asList(contentStyle));


                /*ComplexCellStyle style = new ComplexCellStyle(2, column.getIndex(), IndexedColors.YELLOW.getIndex());
                style.setFontHeightInPoints((short)12);
                style.setBold(true);
                style.setColor(IndexedColors.RED.getIndex());
                style.setFontName("微软雅黑");
                style.setWrapped(true);
                cellStyles.add(style);*/


                ArrayBlockingQueue<ComplexCellStyle> cellStyles = buildCellStyle(sheet);
                ComplexCellStyleStrategy cellStyleStrategy = new ComplexCellStyleStrategy();
                cellStyleStrategy.setCellStyles(cellStyles);

                //列宽设置
                /*ArrayBlockingQueue<ComplexColumnWidthStyle> columnWidthStyles = buildWidthStyle(head);
                ComplexColumnWidthStyleStrategy complexColumnWidthStyleStrategy = new ComplexColumnWidthStyleStrategy();
                complexColumnWidthStyleStrategy.setComplexColumnWidthStyles(columnWidthStyles);*/

                writeSheet.setCustomWriteHandlerList(Arrays.asList(horizontalCellStyleStrategy, cellStyleStrategy));
                excelWriter.write(getData(sheet), writeSheet);
            }

        }
    }

    private ArrayBlockingQueue<ComplexColumnWidthStyle> buildWidthStyle(List<List<String>> head) {
        int size = head.size();
        ArrayBlockingQueue<ComplexColumnWidthStyle> queue = new ArrayBlockingQueue<>(size);
        for (int i = 0; i < size; i++) {
            ComplexColumnWidthStyle widthStyle = new ComplexColumnWidthStyle(26 * 256, i);
            queue.add(widthStyle);
        }
        return queue;
    }

    private List<List<String>> getHead(ExportSheet sheet) {
        List<ExportColumn> columns = sheet.getColumns();
        List<List<String>> result = new ArrayList<>(columns.size());
        for (ExportColumn column : columns) {
            List<String> head = new ArrayList<>();
            head.add(sheet.getSheetCode());
            head.add(column.getColumnCode());
            head.add(column.getColumnName());
            result.add(head);
        }
        return result;
    }


    private List<List<Object>> getData(ExportSheet sheet) {
        List<List<Object>> datas = new ArrayList<>();
        List<ExportColumn> columns = sheet.getColumns();
        for (int i = 0; i < 10; i++) {
            List<Object> data = new ArrayList<>();
            for (int j = 0; j < columns.size(); j++) {
                data.add(i * j);
            }
            datas.add(data);
        }
        return datas;
    }

    private ArrayBlockingQueue<ComplexRowHeightStyle> buildRowHeightStyle(List<List<String>> head) {
        int size = head.get(0).size();
        ArrayBlockingQueue<ComplexRowHeightStyle> result = new ArrayBlockingQueue<>(size);
        for (int i = 0; i < size; i++) {
            ComplexRowHeightStyle complexRowHeightStyle = null;
            if (i == size - 1) {
                complexRowHeightStyle = new ComplexRowHeightStyle((short) 50, (short) 30, i);
            } else {
                complexRowHeightStyle = new ComplexRowHeightStyle((short) 0, (short) 30, i);
            }
            result.add(complexRowHeightStyle);
        }
        return result;
    }

    private ArrayBlockingQueue<ComplexCellStyle> buildCellStyle(ExportSheet sheet) {
        List<ExportColumn> columns = sheet.getColumns();
        ArrayBlockingQueue<ComplexCellStyle> cellStyles = new ArrayBlockingQueue<>(columns.size());
        for (ExportColumn column : columns) {
            if (column.getColumnName().equals("年龄") || column.getColumnName().equals("英语")) {
                ComplexCellStyle style = new ComplexCellStyle(2, column.getIndex(), IndexedColors.YELLOW.getIndex());
                style.setFontHeightInPoints((short) 12);
                style.setBold(true);
                style.setColor(IndexedColors.RED.getIndex());
                style.setFontName("微软雅黑");
                style.setWrapped(true);
                cellStyles.add(style);
            }
        }
        return cellStyles;
    }



}
