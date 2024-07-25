package com.alibaba.easyexcel.test.local.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.fastjson2.JSON;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @Description: 本类用于
 * @Author larry
 * @Date 2023/8/1 10:13
 */
@Data
@NoArgsConstructor
public class HeadsListener extends AnalysisEventListener<Map<Integer, String>> {

    //Excel数据
    private List<Map<Integer, Map<Integer, String>>> list = new ArrayList<>();
    //Excel列名
    private Map<Integer, String> headTitleMap = new HashMap<>();


    private List<Integer> counts;

    public HeadsListener(List<Integer> counts) {
        this.counts = counts;
    }

    private boolean flag = false;


    @Override
    public void invoke(Map<Integer, String> data, AnalysisContext context) {
        if (!flag) {
            counts.add(context.readSheetHolder().getSheetNo());
            flag = true;
        }
        System.out.println("解析到一条数据：" + JSON.toJSONString(data));
        Map<Integer, Map<Integer, String>> map = new HashMap<>();
        map.put(context.readRowHolder().getRowIndex(), data);
        list.add(map);
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        System.out.println("所有数据解析完成");
        flag = false;
    }

    @Override
    public void invokeHeadMap(Map<Integer, String> headMap, AnalysisContext context) {
        headTitleMap = headMap;
    }

}
