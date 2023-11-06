package com.wiblog.poi.util;

import cn.hutool.core.collection.IterUtil;
import cn.hutool.core.exceptions.DependencyException;
import cn.hutool.core.map.MapUtil;
import cn.hutool.core.util.ObjectUtil;
import cn.hutool.poi.PoiChecker;
import cn.hutool.poi.excel.ExcelWriter;
import cn.hutool.poi.excel.cell.CellSetter;
import cn.hutool.poi.excel.cell.setters.CellSetterFactory;
import com.wiblog.poi.excel.reader.PoiExcelReader;
import com.wiblog.poi.excel.writer.PoiExcelWriter;
import com.wiblog.poi.excel.writer.StyleSet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.InputStream;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * @author panwm
 * @since 2023/10/22 15:02
 */
public class PoiExcelUtil {

    /**
     * 获得{@link ExcelWriter}，默认写出到第一个sheet，名字为sheet1
     *
     * @param destFile 目标文件
     * @return {@link ExcelWriter}
     */
    public static PoiExcelWriter getWriter(File destFile) {
        try {
            return new PoiExcelWriter(destFile);
        } catch (NoClassDefFoundError e) {
            throw new DependencyException(ObjectUtil.defaultIfNull(e.getCause(), e), PoiChecker.NO_POI_ERROR_MSG);
        }
    }

    public static PoiExcelReader getReader(File bookFile, int sheetIndex) {
        try {
            return new PoiExcelReader(bookFile, sheetIndex);
        } catch (NoClassDefFoundError e) {
            throw new DependencyException(ObjectUtil.defaultIfNull(e.getCause(), e), PoiChecker.NO_POI_ERROR_MSG);
        }
    }

    public static PoiExcelReader getReader(InputStream bookStream, int sheetIndex) {
        try {
            return new PoiExcelReader(bookStream, sheetIndex);
        } catch (NoClassDefFoundError e) {
            throw new DependencyException(ObjectUtil.defaultIfNull(e.getCause(), e), PoiChecker.NO_POI_ERROR_MSG);
        }
    }

    /**
     * 设置单元格值<br>
     * 根据传入的styleSet自动匹配样式<br>
     * 当为头部样式时默认赋值头部样式，但是头部中如果有数字、日期等类型，将按照数字、日期样式设置
     *
     * @param cell     单元格
     * @param value    值
     * @param styleSet 单元格样式集，包括日期等样式，null表示无样式
     * @param isHeader 是否为标题单元格
     */
    public static void setCellValue(Cell cell, Object value, StyleSet styleSet, boolean isHeader) {
        if (null == cell) {
            return;
        }

        if (null != styleSet) {
            cell.setCellStyle(styleSet.getStyleByValueType(value, isHeader));
        }

        setCellValue(cell, value);
    }

    /**
     * 设置单元格值<br>
     * 根据传入的styleSet自动匹配样式<br>
     * 当为头部样式时默认赋值头部样式，但是头部中如果有数字、日期等类型，将按照数字、日期样式设置
     *
     * @param cell  单元格
     * @param value 值或{@link CellSetter}
     * @since 5.6.4
     */
    public static void setCellValue(Cell cell, Object value) {
        if (null == cell) {
            return;
        }

        // issue#1659@Github
        // 在使用BigWriter(SXSSF)模式写出数据时，单元格值为直接值，非引用值（is标签）
        // 而再使用ExcelWriter(XSSF)编辑时，会写出引用值，导致失效。
        // 此处做法是先清空单元格值，再写入
        if(CellType.BLANK != cell.getCellType()){
            cell.setBlank();
        }

        CellSetterFactory.createCellSetter(value).setValue(cell);
    }

    /**
     * 写一行数据
     *
     * @param row      行
     * @param rowData  一行的数据
     * @param styleSet 单元格样式集，包括日期等样式，null表示无样式
     * @param isHeader 是否为标题行
     */
    public static void writeRow(Row row, Iterable<?> rowData, StyleSet styleSet, boolean isHeader) {
        int i = 0;
        Cell cell;
        for (Object value : rowData) {
            cell = row.createCell(i);
            setCellValue(cell, value, styleSet, isHeader);
            i++;
        }
    }
}
