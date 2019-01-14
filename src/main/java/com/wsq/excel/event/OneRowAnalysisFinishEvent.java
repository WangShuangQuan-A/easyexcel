package com.wsq.excel.event;

import java.util.ArrayList;
import java.util.List;

/**
 * @author jipengfei
 */
public class OneRowAnalysisFinishEvent {

    public OneRowAnalysisFinishEvent(List<String> content) {
        this.data = content;
    }

    public OneRowAnalysisFinishEvent(String[] content, int length) {
        if (content != null) {
            List<String> ls = new ArrayList<String>(length);
            for (int i = 0; i <= length; i++) {
                ls.add(content[i]);
            }
            data = ls;
        }
    }

    private List<String> data;

    public List<String> getData() {
        return data;
    }

    public void setData(List<String> data) {
        this.data = data;
    }
}
