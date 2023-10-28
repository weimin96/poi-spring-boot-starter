package com.wiblog.poi.bean;

import com.wiblog.poi.excel.annotation.Excel;
import com.wiblog.poi.excel.annotation.ExcelHead;
import lombok.Data;

import java.math.RoundingMode;
import java.util.Date;

/**
 * @author panwm
 * @since 2023/10/17 22:20
 */
@Data
@ExcelHead(freezeRow = 1, height = 14)
public class MergeScore {

    private String test1;

    @Excel(name = "班级", replace = {"一年级(1)班", "1班", "一年级(2)班", "2班"}, width = 20)
    private String className;

    @Excel(name = "姓名", sort = 2)
    private String name;

    private String test2;

    @Excel(name = "成绩", merge = true)
    private Score ss;

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

    @Data
    static class Score {

        @Excel(name = "类别", merge = true)
        private Category score;

        private String test;

        @Excel(name = "平均", numFormat = "0.00", roundingMode = RoundingMode.CEILING)
        private Double average;

        @Data
        static class Category {

            private String test;

            @Excel(name = "科目", dictType = "subject", sort = 1)
            private String subject;

            @Excel(name = "分数")
            private Double score;
        }

    }
}
