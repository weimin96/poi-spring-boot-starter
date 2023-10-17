package com.wiblog.poi.bean;

import com.wiblog.poi.excel.annotation.Excel;
import lombok.Data;

/**
 * @author panwm
 * @since 2023/10/17 22:20
 */
@Data
public class Score {

    private String className;

    @Excel(name = "姓名")
    private String name;

    private String subject;

    private Double score;
}
