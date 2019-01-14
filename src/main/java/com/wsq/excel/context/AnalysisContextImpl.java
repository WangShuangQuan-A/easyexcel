package com.wsq.excel.context;

import com.wsq.excel.event.AnalysisEventListener;
import com.wsq.excel.exception.ExcelAnalysisException;
import com.wsq.excel.metadata.BaseRowModel;
import com.wsq.excel.metadata.ExcelColumnProperty;
import com.wsq.excel.metadata.ExcelHeadProperty;
import com.wsq.excel.metadata.Sheet;
import com.wsq.excel.support.ExcelTypeEnum;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 *
 * @author jipengfei
 */
public class AnalysisContextImpl implements AnalysisContext {

    private Object custom;

    private Sheet currentSheet;

    private ExcelTypeEnum excelType;

    private InputStream inputStream;

    private AnalysisEventListener eventListener;

    private Integer currentRowNum;

    private Integer totalCount;

    private ExcelHeadProperty excelHeadProperty;

    private boolean trim;

    private boolean use1904WindowDate = false;

    @Override
    public void setUse1904WindowDate(boolean use1904WindowDate) {
        this.use1904WindowDate = use1904WindowDate;
    }

    @Override
    public Object getCurrentRowAnalysisResult() {
        return currentRowAnalysisResult;
    }

    @Override
    public void interrupt() {
        throw new ExcelAnalysisException("interrupt error");
    }

    @Override
    public boolean use1904WindowDate() {
        return use1904WindowDate;
    }

    @Override
    public void setCurrentRowAnalysisResult(Object currentRowAnalysisResult) {
        this.currentRowAnalysisResult = currentRowAnalysisResult;
    }

    private Object currentRowAnalysisResult;

    public AnalysisContextImpl(InputStream inputStream, ExcelTypeEnum excelTypeEnum, Object custom,
                               AnalysisEventListener listener, boolean trim) {
        this.custom = custom;
        this.eventListener = listener;
        this.inputStream = inputStream;
        this.excelType = excelTypeEnum;
        this.trim = trim;
    }

    @Override
    public void setCurrentSheet(Sheet currentSheet) {
        cleanCurrentSheet();
        this.currentSheet = currentSheet;
        if (currentSheet.getClazz() != null) {
            buildExcelHeadProperty(currentSheet.getClazz(), null);
        }
    }

    private void cleanCurrentSheet() {
        this.currentSheet = null;
        this.excelHeadProperty = null;
        this.totalCount = 0;
        this.currentRowAnalysisResult = null;
        this.currentRowNum =0;
    }

    @Override
    public ExcelTypeEnum getExcelType() {
        return excelType;
    }

    public void setExcelType(ExcelTypeEnum excelType) {
        this.excelType = excelType;
    }

    @Override
    public Object getCustom() {
        return custom;
    }

    public void setCustom(Object custom) {
        this.custom = custom;
    }

    @Override
    public Sheet getCurrentSheet() {
        return currentSheet;
    }

    @Override
    public InputStream getInputStream() {
        return inputStream;
    }

    public void setInputStream(InputStream inputStream) {
        this.inputStream = inputStream;
    }

    @Override
    public AnalysisEventListener getEventListener() {
        return eventListener;
    }

    public void setEventListener(AnalysisEventListener eventListener) {
        this.eventListener = eventListener;
    }

    @Override
    public Integer getCurrentRowNum() {
        return this.currentRowNum;
    }

    @Override
    public void setCurrentRowNum(Integer row) {
        this.currentRowNum = row;
    }

    @Override
    public Integer getTotalCount() {
        return totalCount;
    }

    @Override
    public void setTotalCount(Integer totalCount) {
        this.totalCount = totalCount;
    }

    @Override
    public ExcelHeadProperty getExcelHeadProperty() {
        return this.excelHeadProperty;
    }

    @Override
    public void buildExcelHeadProperty(Class<? extends BaseRowModel> clazz, List<String> headOneRow) {
        if (this.excelHeadProperty == null && (clazz != null || headOneRow != null)) {
            this.excelHeadProperty = new ExcelHeadProperty(clazz, new ArrayList<>());
        }
        List<String> head = this.excelHeadProperty.getHead();
        if (head == null && headOneRow != null) {
            this.excelHeadProperty.appendOneRow(headOneRow);
        }
        if (head != null && headOneRow != null) {
            for (int i = 0; i < headOneRow.size(); i++) {
                for (String s : head) {
                    if (headOneRow.get(i).equals(s)) {
                        ExcelColumnProperty excelColumnProperty = this.excelHeadProperty.getExcelColumnPropertyMap2().get(s);
                        this.excelHeadProperty.getExcelColumnPropertyMap1().put(i,excelColumnProperty);
                    }
                }
            }
        }
    }

    @Override
    public boolean trim() {
        return this.trim;
    }
}
