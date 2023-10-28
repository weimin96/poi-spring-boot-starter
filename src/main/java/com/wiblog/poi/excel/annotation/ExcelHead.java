package com.wiblog.poi.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author pwm
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface ExcelHead {

    /**
     * 冻结行数
     */
    int freezeRow() default 0;

    /**
     * 导出时在excel中每个列的高度 单位为字符
     */
    int height() default 14;
}
