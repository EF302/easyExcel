package com.lwl.easyexcel.listener;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.exception.ExcelDataConvertException;
import com.alibaba.fastjson.JSON;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;

/**
 * 监听器
 */
@Slf4j
public class ExcelReadListener<T> extends AnalysisEventListener<T> {

    /**
     * 批处理阈值（每次读取多少数据后，就进行存储）
     * 注 正是因为可以设置读取处理数据的阈值，所以easyExcel才能有效避免把超大量数据同时加载到内存的情况
     */
    private static final int BATCH_COUNT = 2;

    /**
     * 临时存储读取excel数据的集合
     */
    private List<T> dataList = new ArrayList<>(BATCH_COUNT);

    /**
     * 读取表头信息
     *
     * @param headMap
     * @param context
     */
    @Override
    public void invokeHeadMap(Map<Integer, String> headMap, AnalysisContext context) {
        log.info("解析到一条头数据:{}", JSON.toJSONString(headMap));
    }

    /**
     * 每一条数据解析都会来调用
     *
     * @param user
     * @param analysisContext
     */
    @Override
    public void invoke(T user, AnalysisContext analysisContext) {
        log.info("解析到一条数据:{}", JSON.toJSONString(user));
        dataList.add(user);
        if (dataList.size() >= BATCH_COUNT) {
            log.info("读取数据达到指定阈值，开始进行数据存储！");
            this.saveData();
            dataList.clear();
        }
    }

    /**
     * 读取完数据后的动作：用于处理最后一点没有存储的数据
     *
     * @param analysisContext
     */
    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        if (CollectionUtils.isNotEmpty(dataList)) {
            log.info("数据读取完毕，处理末余数据！");
            this.saveData();
            System.out.println(dataList);
        }
        log.info("所有数据解析完成！");
    }

    /**
     * 异常信息处理
     */
    @Override
    public void onException(Exception exception, AnalysisContext context) {
        log.error("解析失败，但是继续解析下一行:{}", exception.getMessage());
        // 如果是某一个单元格的转换异常，能获取到具体行号，如果要获取头的信息，配合invokeHeadMap使用
        if (exception instanceof ExcelDataConvertException) {
            ExcelDataConvertException excelDataConvertException = (ExcelDataConvertException) exception;
            log.error("第{}行，第{}列，数据：{}，解析异常", excelDataConvertException.getRowIndex(), excelDataConvertException.getColumnIndex(), excelDataConvertException.getCellData());
        }
    }

    /**
     * 模拟数据存储
     */
    private void saveData() {
        log.info("{}条数据，开始存储数据库！", dataList.size());
        log.info("存储数据库成功！");
    }
}
