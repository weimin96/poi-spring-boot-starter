package com.wiblog.poi.excel.reader;

import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import cn.hutool.poi.excel.StyleSet;
import com.wiblog.poi.excel.ExportParam;
import com.wiblog.poi.excel.annotation.Excel;
import com.wiblog.poi.excel.annotation.ExcelHead;
import com.wiblog.poi.excel.handler.IExcelDictHandler;
import com.wiblog.poi.util.PoiExcelUtil;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.util.StringUtil;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.util.Date;
import java.util.List;

/**
 * @author panwm
 * @since 2023/10/17 22:47
 */
public class ExcelHandler {

    private final IExcelDictHandler dictHandler;

    public ExcelHandler(IExcelDictHandler dictHandler) {
        this.dictHandler = dictHandler;
    }

    public <T> List<T> readExcel(File excel, Class<T> beanType) {
        return readExcel(excel, 0, 0, 1, beanType);
    }

    public <T> List<T> readExcel(InputStream inputStream, Class<T> beanType) {
        return readExcel(inputStream, 0, 0, 1, beanType);
    }

    public <T> List<T> readExcel(InputStream inputStream, int sheetIndex, int headerRowIndex, int startRowIndex, Class<T> beanType) {
        PoiExcelReader reader = PoiExcelUtil.getReader(inputStream, sheetIndex);
        return readExcel(reader, headerRowIndex, startRowIndex, beanType);
    }

    public <T> List<T> readExcel(File file, int sheetIndex, int headerRowIndex, int startRowIndex, Class<T> beanType) {
        PoiExcelReader reader = PoiExcelUtil.getReader(file, sheetIndex);
        return readExcel(reader, headerRowIndex, startRowIndex, beanType);
    }

    /**
     * @param reader         reader
     * @param headerRowIndex – 标题所在行，如果标题行在读取的内容行中间，这行做为数据将忽略，从0开始计数
     * @param startRowIndex  – 起始行（包含，从0开始计数）
     * @param beanType       实体类型
     * @param <T>
     * @return
     */
    public <T> List<T> readExcel(PoiExcelReader reader, int headerRowIndex, int startRowIndex, Class<T> beanType) {
        return reader.read(headerRowIndex, startRowIndex, Integer.MAX_VALUE, beanType, dictHandler);
    }

    public void writeExcel(ExportParam param, File file) {
        ExcelWriter writer = ExcelUtil.getBigWriter(file);
        writeData(param, writer);
        writer.write(param.getData(), true);
        // 关闭writer，释放内存
        writer.close();
    }

    public void writeExcel(HttpServletResponse response, ExportParam param) {
        // 在内存操作，写出到浏览器
        ExcelWriter writer = ExcelUtil.getBigWriter();
        writeData(param, writer);
        // 设置浏览器响应的格式
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8");
        String fileName = null;
        try {
            fileName = URLEncoder.encode(param.getTitle(), "UTF-8");
            response.setHeader("Content-Disposition", "attachment;filename=" + fileName + ".xlsx");

            ServletOutputStream out = response.getOutputStream();
            writer.flush(out, true);
            out.close();
            writer.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private void writeData(ExportParam param, ExcelWriter writer) {
        // 获取类注解
        Class<?> clazz = param.getEntity();
        if (clazz.isAnnotationPresent(ExcelHead.class)) {
            ExcelHead annotation = clazz.getAnnotation(ExcelHead.class);
            int i = annotation.freezeRow();
            if (i > 0) {
                writer.setFreezePane(i);
            }
            int height = annotation.height();
            writer.setDefaultRowHeight(height);
        }

        // 获取所有字段
        Field[] fields = clazz.getDeclaredFields();
        Sheet sheet = writer.getSheet();
        for (int i = 0; i < fields.length; i++) {
            Field field = fields[i];
            if (field.isAnnotationPresent(Excel.class)) {
                Excel annotation = field.getAnnotation(Excel.class);
                // 设置名称
                String name = annotation.name();
                if ("".equals(name)) {
                    name = field.getName();
                }
                writer.addHeaderAlias(field.getName(), name);

                // 设置宽度
                int width = annotation.width();
                if (width > 0) {
                    writer.setColumnWidth(i, width);
                }

                // 设置时间格式转换
                String dateFormat = annotation.dateFormat();
                if (StringUtil.isNotBlank(dateFormat) && field.getType().equals(Date.class)) {
                    DataFormat dataFormat = writer.getWorkbook().createDataFormat();
                    short format = dataFormat.getFormat(dateFormat);
                    StyleSet styleSet = writer.getStyleSet();
                    styleSet.getCellStyleForDate()
                            .setDataFormat(format);
                }

                // 设置翻译

                String dictType = annotation.dictType();
                if (StringUtil.isNotBlank(dictType)) {
                }
            }
        }

        //排除字段操作
        writer.setOnlyAlias(true);

        //设置sheet的名称
        if (StringUtil.isNotBlank(param.getSheetName())) {
            writer.renameSheet(param.getSheetName());
        }
    }

}
