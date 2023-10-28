package com.wiblog.poi.excel.handler;

import java.util.List;
import java.util.Map;

/**
 * @author panwm
 * @since 2023/10/21 23:01
 */
public interface IExcelDictHandler {

    /**
     * 返回字典所有值
     * key: dictKey
     * value: dictValue
     * @param dict  字典Key
     * @return List
     */
    default List getList(String dict){return null;}

    /**
     * 从值翻译到名称
     *
     * @param dict  字典Key
     * @param value 属性值
     * @return
     */
    String toName(String dict, Object value);

    /**
     * 从名称翻译到值
     *
     * @param dict  字典Key
     * @param name  属性名称
     * @return
     */
    String toValue(String dict,String name);
}
