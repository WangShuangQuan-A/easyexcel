package com.wsq.excel.modelbuild;

import com.wsq.excel.context.AnalysisContext;
import com.wsq.excel.event.AnalysisEventListener;
import com.wsq.excel.exception.ExcelGenerateException;
import com.wsq.excel.metadata.ExcelHeadProperty;
import com.wsq.excel.util.TypeUtil;
import net.sf.cglib.beans.BeanMap;

import java.util.List;
import java.util.Map;

/**
 * @author jipengfei
 */
public class ModelBuildEventListener implements AnalysisEventListener {

    @Override
    public void invoke(Object object, AnalysisContext context) {
        if (context.getExcelHeadProperty() != null && context.getExcelHeadProperty().getHeadClazz() != null) {
            try {
                Object resultModel = buildUserModel(context, (List<String>)object);
                context.setCurrentRowAnalysisResult(resultModel);
            } catch (Exception e) {
                throw new ExcelGenerateException(e);
            }
        }
    }

    @Override
    public boolean doAfterAllAnalysed(AnalysisContext context) {
        return true;
    }

    private Object buildUserModel(AnalysisContext context, List<String> stringList) throws Exception {
        ExcelHeadProperty excelHeadProperty = context.getExcelHeadProperty();
        Object resultModel = excelHeadProperty.getHeadClazz().newInstance();
        BeanMap.create(resultModel).putAll(
            TypeUtil.getFieldValues(stringList, excelHeadProperty, context.use1904WindowDate()));
        return resultModel;
    }

/*    @Override
    public Map<String, Object> getParams() {
        return null;
    }

    @Override
    public void setParams(Map map) {

    }*/
}
