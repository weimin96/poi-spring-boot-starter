package com.wiblog.poi.excel;

import lombok.Builder;
import lombok.Data;
import lombok.experimental.Accessors;

import java.util.List;

/**
 * @author panwm
 * @since 2023/10/19 0:35
 */
@Builder
@Accessors(chain = true)
@Data
public class ExportParam {

    /**
     * 导出数据
     */
    private Iterable<?> data;

//    private ExportParams title;

    private String title;

    /**
     * 导出实体
     */
    private Class<?> entity;

    /**
     * 导出sheet名字
     */
    private String sheetName;
}
