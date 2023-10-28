package com.wiblog.poi.bean;

import com.wiblog.poi.excel.annotation.Excel;
import com.wiblog.poi.excel.annotation.ExcelHead;
import lombok.Data;

import java.util.Date;

/**
 * @author panwm
 * @since 2023/10/17 22:20
 */
@Data
@ExcelHead(freezeRow = 1, height = 14)
public class Score {

    @Excel(name = "班级", width = 20)
    private String className;

    @Excel(name = "姓名", sort = 2)
    private String name;

    private String test1;

    @Excel(name = "科目", dictType = "subject", sort = 1)
    private String subject;

    @Excel(name = "成绩")
    private Double score;

    @Excel(name = "年份")
    private Date year;

    @Excel(name = "日期", dateFormat = "yyyy年MM月dd日", width = 30)
    private Date date;

    private String test;

    @Excel(name = "性别", dictType = "sex")
    private String sex;

    @Excel(name = "备注1", defaultValue = "默认备注1")
    private String remark1;

    @Excel(name = "备注2")
    private String remark2;
}
