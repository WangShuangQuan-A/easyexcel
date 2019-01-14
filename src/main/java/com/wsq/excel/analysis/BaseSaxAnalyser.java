package com.wsq.excel.analysis;

import com.wsq.excel.context.AnalysisContext;
import com.wsq.excel.event.AnalysisEventListener;
import com.wsq.excel.event.AnalysisEventRegisterCenter;
import com.wsq.excel.event.OneRowAnalysisFinishEvent;
import com.wsq.excel.metadata.Sheet;
import com.wsq.excel.util.TypeUtil;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * @author jipengfei
 */
public abstract class BaseSaxAnalyser implements AnalysisEventRegisterCenter, ExcelAnalyser {

    protected AnalysisContext analysisContext;

    private LinkedHashMap<String, AnalysisEventListener> listeners = new LinkedHashMap<String, AnalysisEventListener>();

    /**
     * execute method
     */
    protected abstract void execute();


    @Override
    public void appendLister(String name, AnalysisEventListener listener) {
        if (!listeners.containsKey(name)) {
            listeners.put(name, listener);
        }
    }

    @Override
    public void analysis(Sheet sheetParam) {
        execute();
    }

    @Override
    public void analysis() {
        execute();
    }

    /**
     */
    @Override
    public void cleanAllListeners() {
        listeners = new LinkedHashMap<String, AnalysisEventListener>();
    }

    @Override
    public void notifyListeners(OneRowAnalysisFinishEvent event) {
        analysisContext.setCurrentRowAnalysisResult(event.getData());
        /** Parsing header content **/
        if (analysisContext.getCurrentRowNum() < analysisContext.getCurrentSheet().getHeadLineMun()) {
            if (analysisContext.getCurrentRowNum() <= analysisContext.getCurrentSheet().getHeadLineMun() - 1) {
                analysisContext.buildExcelHeadProperty(analysisContext.getCurrentSheet().getClazz(),
                    (List<String>)analysisContext.getCurrentRowAnalysisResult());
            }
        } else {
            /**  notify all event listeners **/
            for (Map.Entry<String, AnalysisEventListener> entry : listeners.entrySet()) {
                entry.getValue().invoke(analysisContext.getCurrentRowAnalysisResult(), analysisContext);
            }
        }
    }

    private List<String> converter(List<String> data) {
        List<String> list = new ArrayList<String>();
        if (data != null) {
            for (String str : data) {
                list.add(TypeUtil.formatFloat(str));
            }
        }
        return list;
    }

}
