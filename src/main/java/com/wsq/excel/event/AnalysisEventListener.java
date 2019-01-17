package com.wsq.excel.event;

import com.wsq.excel.context.AnalysisContext;

import java.util.Map;

/**
 * @author wsq
 */
public interface AnalysisEventListener<T> {

    /**
     * when analysis one row trigger invoke function
     *
     * @param object  one row data
     * @param context analysis context
     */
    void invoke(T object, AnalysisContext context);

    /**
     * if have something to do after all  analysis
     */
    void doAfterAllAnalysed(AnalysisContext context);

    /**
     * 扩展字段
     */
    Map<String,Object> getParams();

    void setParams(Map<String,Object> params);
}
