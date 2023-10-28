package com.wiblog.poi.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
import java.math.RoundingMode;

/**
 * @author pwm
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface Excel {

    /**
     * 导出到Excel中的名字.
     */
    String name() default "";

    /**
     * 导入导出日期格式化, 如果是String类型，则导入时生效，如果是Date类型，则导出时生效，如: yyyy-MM-dd
     */
    String dateFormat() default "";

    /**
     * 导入导出数值格式化 例：保留两位小数：0.00
     */
    String numFormat() default "";

    /**
     * numFormat 舍入规则 默认:RoundingMode.HALF_UP 四舍五入（RoundingMode.FLOOR：向下取整，RoundingMode.CEILING 向上取整）
     */
    RoundingMode roundingMode() default RoundingMode.HALF_UP;

    /**
     * 导出时在excel中每个列的宽 单位为字符
     */
    int width() default 0;

    /**
     * 导出时在excel中排序
     */
    int sort() default Integer.MAX_VALUE;

    /**
     * 如果是字典类型，请设置字典的type值 (如: sys_user_sex)
     */
    String dictType() default "";

    /**
     * 导入数据时是否需要处理纵向合并，用于Bean类型
     */
    boolean merge() default false;

    /**
     * 值替换 {a,b,...}：导出{a -> b，...} ，导入 {a <- b,...}
     */
    String[] replace() default {};


    /**
     * 导出类型（0数字 1字符串）
     */
    ColumnType cellType() default ColumnType.STRING;


    /**
     * 文字后缀,如% 90 变成90%
     */
    String suffix() default "";

    /**
     * 当值为空时,字段的默认值
     */
    String defaultValue() default "";

    /**
     * 提示信息
     */
    String prompt() default "";

    /**
     * 设置只能选择不能输入的列内容.
     */
    String[] combo() default {};

    /**
     * 是否导出数据,应对需求:有时我们需要导出一份模板,这是标题需要但内容需要用户手工填写.
     */
    boolean isExport() default true;

    /**
     * 另一个类中的属性名称,支持多级获取,以小数点隔开
     */
    String targetAttr() default "";

    /**
     * 是否自动统计数据,在最后追加一行统计数据总和
     */
    boolean isStatistics() default false;

    /**
     * 导出字段对齐方式（0：默认；1：靠左；2：居中；3：靠右）
     */
    Align align() default Align.AUTO;

    enum Align {
        AUTO(0), LEFT(1), CENTER(2), RIGHT(3);
        private final int value;

        Align(int value) {
            this.value = value;
        }

        public int value() {
            return this.value;
        }
    }

    /**
     * 字段类型（0：导出导入；1：仅导出；2：仅导入）
     */
    Type type() default Type.ALL;

    enum Type {
        ALL(0), EXPORT(1), IMPORT(2);
        private final int value;

        Type(int value) {
            this.value = value;
        }

        public int value() {
            return this.value;
        }
    }

    enum ColumnType {
        NUMERIC(0), STRING(1), IMAGE(2);
        private final int value;

        ColumnType(int value) {
            this.value = value;
        }

        public int value() {
            return this.value;
        }
    }
}
