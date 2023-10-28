package com.wiblog.poi.service;

import com.wiblog.poi.bean.DictData;
import com.wiblog.poi.excel.handler.IExcelDictHandler;
import org.assertj.core.util.Maps;
import org.springframework.stereotype.Component;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author panwm
 * @since 2023/10/21 17:08
 */
@Component
public class DictService {

    public List<DictData> getDictList(String dict) {
        List<DictData> list = new ArrayList<>();
        if ("subject".equals(dict)) {
            list = new ArrayList<>();
            DictData data11 = new DictData();
            data11.setLabel("语文");
            data11.setValue("1");

            DictData data12 = new DictData();
            data12.setLabel("数学");
            data12.setValue("2");

            list.add(data11);
            list.add(data12);
        } else if ("sex".equals(dict)) {
            DictData data21 = new DictData();
            data21.setLabel("男");
            data21.setValue("1");

            DictData data22 = new DictData();
            data22.setLabel("女");
            data22.setValue("2");

            list.add(data21);
            list.add(data22);
        }
        return list;
    }

}
