package com.wiblog.poi.excel.bean;

import lombok.Data;

import java.util.Map;

/**
 * @author panwm
 * @since 2023/10/22 14:39
 */
@Data
public class MergeCell {

    private int index;

    private String field;

    private String name;

    private Map<String, MergeCell> children;

//    private Class clazz;

}
