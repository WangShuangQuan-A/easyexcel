package com.wsq.easyexcel.test.listen;

import com.wsq.excel.context.AnalysisContext;
import com.wsq.excel.event.AnalysisEventListener;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class ExcelListener implements AnalysisEventListener {


    private List<Object>  data = new ArrayList<Object>();

    @Override
    public void invoke(Object object, AnalysisContext context) {
        System.out.println(context.getCurrentSheet());
        data.add(object);
        if(data.size()>=100){
            doSomething();
            data = new ArrayList<Object>();
        }
    }

    @Override
    public boolean doAfterAllAnalysed(AnalysisContext context) {
        return true;//doSomething();
    }

/*    @Override
    public Map<String, Object> getParams() {
        return null;
    }

    @Override
    public void setParams(Map map) {

    }*/

    public void doSomething(){
        for (Object o:data) {
            System.out.println(o);
        }
    }
}
