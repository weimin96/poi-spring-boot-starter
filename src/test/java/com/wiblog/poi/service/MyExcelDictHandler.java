package com.wiblog.poi.service;

import cn.hutool.core.util.ObjectUtil;
import com.wiblog.poi.bean.DictData;
import com.wiblog.poi.excel.handler.IExcelDictHandler;
import org.junit.platform.commons.util.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author panwm
 * @since 2023/10/21 17:08
 */
@Component
public class MyExcelDictHandler implements IExcelDictHandler {

    @Autowired
    private DictService dictService;

    private final Map<String, List<DictData>> DICT_MAPS = new HashMap<>(32);

    @Override
    public List getList(String dict) {
        List<DictData> dictList = DICT_MAPS.get(dict);
        if (dictList == null || dictList.isEmpty()) {
            dictList = dictService.getDictList(dict);
            DICT_MAPS.put(dict, dictList);
        }
        return dictList;
    }

    @Override
    public String toName(String dict, Object value) {
        List<DictData> list = getList(dict);
        if (list.isEmpty()) {
            return null;
        }
        String key = ObjectUtil.toString(value).trim();
        for (DictData data : list) {
            if (key.equals(data.getValue())) {
                return data.getLabel();
            }
        }
        return null;
    }

    @Override
    public String toValue(String dict, String name) {
        if (StringUtils.isBlank(dict) || StringUtils.isBlank(name)) {
            return null;
        }
        List<DictData> list = getList(dict);
        if (list.isEmpty()) {
            return null;
        }
        name = name.trim();
        for (DictData data : list) {
            if (name.equals(data.getLabel())) {
                return data.getValue();
            }
        }
        return null;
    }


}
