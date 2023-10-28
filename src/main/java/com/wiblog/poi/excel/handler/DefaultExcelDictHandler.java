package com.wiblog.poi.excel.handler;

import cn.hutool.core.util.ObjectUtil;

/**
 * @author panwm
 * @since 2023/10/21 23:14
 */
public class DefaultExcelDictHandler implements IExcelDictHandler{


    @Override
    public String toName(String dict, Object value) {
        return ObjectUtil.toString(value);
    }

    @Override
    public String toValue(String dict, String name) {
        return name;
    }
}
