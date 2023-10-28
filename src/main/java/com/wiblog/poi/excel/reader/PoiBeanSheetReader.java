package com.wiblog.poi.excel.reader;

import cn.hutool.core.bean.BeanUtil;
import cn.hutool.core.bean.copier.CopyOptions;
import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.collection.IterUtil;
import cn.hutool.core.collection.ListUtil;
import cn.hutool.core.date.DateUtil;
import cn.hutool.core.util.ObjectUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.reader.AbstractSheetReader;
import com.wiblog.poi.excel.annotation.Excel;
import com.wiblog.poi.excel.bean.MergeCell;
import com.wiblog.poi.excel.handler.IExcelDictHandler;
import com.wiblog.poi.exception.ExcelErrorException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.util.StringUtil;

import java.lang.reflect.Field;
import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.util.*;

/**
 * @author panwm
 * @since 2023/10/22 15:14
 */
public class PoiBeanSheetReader<T> extends AbstractSheetReader<List<T>> {

    private final Class<T> beanClass;

    private final int headerRowIndex;

    private final int startRowIndex;

    private final int endRowIndex;

    private IExcelDictHandler dictHandler;

    private final Map<String, String> headerAlias;

    private final Map<String, MergeCell> mergeCellMap;

    private final Map<String, String> dictTypeMap;

    private final Map<String, Map<String, String>> replaceMap;

    private final Map<String, String> dateFormatMap;

    private final Map<String, Map<String,Object>> numFormatMap;

    /**
     * 构造
     *
     * @param headerRowIndex 标题所在行，如果标题行在读取的内容行中间，这行做为数据将忽略
     * @param startRowIndex  起始行（包含，从0开始计数）
     * @param endRowIndex    结束行（包含，从0开始计数）
     * @param beanClass      每行对应Bean的类型
     */
    public PoiBeanSheetReader(int headerRowIndex, int startRowIndex, int endRowIndex, Class<T> beanClass, IExcelDictHandler dictHandler) {
        super(startRowIndex, endRowIndex);
        this.beanClass = beanClass;
        this.headerRowIndex = headerRowIndex;
        this.startRowIndex = startRowIndex;
        this.endRowIndex = endRowIndex;
        this.dictHandler = dictHandler;
        this.headerAlias = new HashMap<>(32);
        this.mergeCellMap = new HashMap<>(32);
        this.dictTypeMap = new HashMap<>(32);
        this.replaceMap = new HashMap<>(32);
        this.dateFormatMap = new HashMap<>(32);
        this.numFormatMap = new HashMap<>(32);
    }

    @Override
    public List<T> read(Sheet sheet) {

        // 获取所有字段
        Field[] fields = this.beanClass.getDeclaredFields();
        List<String> mergeList = new ArrayList<>();
        for (Field field : fields) {
            if (field.isAnnotationPresent(Excel.class)) {
                Excel annotation = field.getAnnotation(Excel.class);
                // 设置名称
                String name = annotation.name();
                if ("".equals(name)) {
                    name = field.getName();
                }
                // 添加名称映射
                addHeaderAlias(name, field.getName());
                // 添加合并单元格映射
                setMergeCell(field);
                // 值替换
                setReplaceMap(annotation, field);
                // 时间格式化
                setDateFormatMap(annotation, field);
                // 数值格式化
                setNumFormatMap(annotation, field);
                // 字典翻译
                String dictType = annotation.dictType();
                if (StringUtil.isNotBlank(dictType)) {
                    dictTypeMap.put(field.getName(), dictType);
                }
                if (annotation.merge()) {
                    mergeList.add(name);
                }
            }
        }

        // 边界判断
        final int firstRowNum = sheet.getFirstRowNum();
        final int lastRowNum = sheet.getLastRowNum();
        if (lastRowNum < 0) {
            return ListUtil.empty();
        }

        if (headerRowIndex < firstRowNum) {
            throw new IndexOutOfBoundsException(StrUtil.format("Header row index {} is lower than first row index {}.", headerRowIndex, firstRowNum));
        } else if (headerRowIndex > lastRowNum) {
            throw new IndexOutOfBoundsException(StrUtil.format("Header row index {} is greater than last row index {}.", headerRowIndex, lastRowNum));
        } else if (startRowIndex > lastRowNum) {
            return ListUtil.empty();
        }
        final int startRowIndex = Math.max(this.startRowIndex, firstRowNum);// 读取起始行（包含）
        final int endRowIndex = Math.min(this.endRowIndex, lastRowNum);// 读取结束行（包含）

        // 读取header
        final List<String> headerList = aliasHeader(readRow(sheet, startRowIndex - 1));
        final List<T> beanList = new ArrayList<>(endRowIndex - startRowIndex + 1);
        final CopyOptions copyOptions = CopyOptions.create().setIgnoreError(true);
        List<Object> rowList;
        for (int i = startRowIndex; i <= endRowIndex; i++) {
            // 跳过标题行
            if (i != headerRowIndex) {
                rowList = readRow(sheet, i);
                if (CollUtil.isNotEmpty(rowList) || !ignoreEmptyRow) {

                    Map<String, Object> dataMap = IterUtil.toMap(headerList, rowList, true);
                    // 值替换
                    handlerReplace(dataMap);
                    // 设置翻译
                    handlerTranslators(dataMap);
                    // 设置日期格式
                    handlerDateFormat(dataMap);
                    // 设置数值格式化
                    handlerNumFormat(dataMap);
                    // 合并单元格处理
                    for (String mergeName : mergeList) {
                        initializeFields(this.mergeCellMap.get(mergeName), dataMap, 0, new HashMap<>());
                    }

                    T bean = BeanUtil.toBean(dataMap, this.beanClass, copyOptions);

                    beanList.add(bean);
                }
            }
        }
        return beanList;
    }

    public void handlerReplace(Map<String, Object> dataMap) {
        for (String filedName : dataMap.keySet()) {
            Map<String, String> repMap = replaceMap.get(filedName);
            if (repMap != null) {
                String val = ObjectUtil.toString(dataMap.get(filedName));
                for (String key: repMap.keySet()) {
                    if (key.equals(val)) {
                        dataMap.put(filedName, repMap.get(key));
                    }
                }
            }
        }
    }

    public void handlerTranslators(Map<String, Object> dataMap) {
        for (String filedName : dataMap.keySet()) {
            String dictType = dictTypeMap.get(filedName);
            if (StringUtil.isNotBlank(dictType)) {
                dataMap.put(filedName, dictHandler.toValue(dictType, ObjectUtil.toString(dataMap.get(filedName))));
            }
        }
    }

    public void handlerDateFormat(Map<String, Object> dataMap) {
        for (String filedName : dataMap.keySet()) {
            String dataFormat = dateFormatMap.get(filedName);
            if (StringUtil.isNotBlank(dataFormat)) {
                try {
                    Date date = (Date) dataMap.get(filedName);
                    dataMap.put(filedName, DateUtil.format(date, dataFormat));
                } catch (Exception ignored) {
                }
            }
        }
    }

    public void handlerNumFormat(Map<String, Object> dataMap) {
        for (String filedName : dataMap.keySet()) {
            Map<String, Object> formatMap = numFormatMap.get(filedName);

            if (formatMap != null) {
                String numFormat = (String) formatMap.get("numFormat");
                RoundingMode round = (RoundingMode) formatMap.get("round");
                Object o = dataMap.get(filedName);
                if (o == null) {
                    continue;
                }
                DecimalFormat df = new DecimalFormat(numFormat);
                df.setRoundingMode(round);
                dataMap.put(filedName, df.format(o));
            }
        }
    }

    public void setReplaceMap(Excel annotation, Field field) {
        String[] replaces = annotation.replace();
        if (replaces.length > 0) {
            if (replaces.length % 2 != 0) {
                throw new ExcelErrorException("replace格式错误，长度必须是2的整数倍");
            }
            HashMap<String, String> rep = new HashMap<>(8);
            this.replaceMap.put(field.getName(), rep);
            for (int i = 0; i < replaces.length; i = i + 2) {
                rep.put(replaces[i + 1], replaces[i]);
            }
        }
    }

    public void setDateFormatMap(Excel annotation, Field field) {
        String dateFormat = annotation.dateFormat();
        if (StringUtil.isNotBlank(dateFormat) && field.getType() == String.class) {
            this.dateFormatMap.put(field.getName(), dateFormat);
        }
    }

    public void setNumFormatMap(Excel annotation, Field field) {
        String numFormat = annotation.numFormat();
        if (StringUtil.isNotBlank(numFormat)) {
            RoundingMode roundingEnum = annotation.roundingMode();
            HashMap<String, Object> map = new HashMap<>(2);
            map.put("numFormat", numFormat);
            map.put("round", roundingEnum);
            this.numFormatMap.put(field.getName(), map);
        }
    }

    /**
     * 合并单元格映射处理
     *
     * @param field field
     */
    public void setMergeCell(Field field) {
        if (!field.isAnnotationPresent(Excel.class)) {
            return;
        }
        Excel annotation = field.getAnnotation(Excel.class);
        // 别名
        addHeaderAlias(annotation.name(), field.getName());
        // 字典
        if (StringUtil.isNotBlank(annotation.dictType())) {
            dictTypeMap.put(field.getName(), annotation.dictType());
        }
        // 值替换
        setReplaceMap(annotation, field);
        // 日期格式化
        setDateFormatMap(annotation, field);
        // 数值格式化
        setNumFormatMap(annotation, field);
        if (!annotation.merge()) {
            return;
        }
        // 合并单元格
        MergeCell mergeCell = new MergeCell();
        mergeCell.setField(field.getName());
        mergeCell.setClazz(field.getType());
        this.mergeCellMap.put(annotation.name(), mergeCell);


        Field[] fields = field.getType().getDeclaredFields();
        for (Field child : fields) {
            setMergeCell(child);
        }
    }

    /**
     * 递归构造合并单元格的类型数据
     *
     * @param mergeCell 类型映射
     * @param dataMap   所有数值映射
     * @param index     递归层级
     * @param fieldMap  传递嵌套类型
     */
    public void initializeFields(MergeCell mergeCell, Map<String, Object> dataMap, int index, Map<String, Object> fieldMap) {
        if (mergeCell == null) {
            return;
        }
        // 获取类的 Class 对象
        Class<?> clazz = mergeCell.getClazz();
        // 获取类的字段
        Field[] fields = clazz.getDeclaredFields();

        if (index == 0) {
            dataMap.put(mergeCell.getField(), fieldMap);
        }

        // 递归初始化属性
        for (Field field : fields) {
            if (!field.isAnnotationPresent(Excel.class)) {
                continue;
            }
            Excel fieldAnnotation = field.getAnnotation(Excel.class);
            String fieldName = field.getName();
            boolean merge = fieldAnnotation.merge();
            Object value = dataMap.get(fieldName);
            // 忽略字段
            if (value == null && !merge) {
                continue;
            }
            // 如果是合并单元格就递归处理
            if (merge) {
                Map<String, Object> childMap = new HashMap<>();
                fieldMap.put(fieldName, childMap);
                initializeFields(this.mergeCellMap.get(fieldAnnotation.name()), dataMap, ++index, childMap);
            } else {
                fieldMap.put(fieldName, value);
            }
        }


    }

    @Override
    public List<String> aliasHeader(List<Object> headerList) {
        if (CollUtil.isEmpty(headerList)) {
            return new ArrayList<>(0);
        }

        final int size = headerList.size();
        final ArrayList<String> result = new ArrayList<>(size);
        for (int i = 0; i < size; i++) {
            result.add(aliasHeader(headerList.get(i), i));
        }
        return result;
    }

    @Override
    protected String aliasHeader(Object headerObj, int index) {
        if (null == headerObj) {
            return ExcelUtil.indexToColName(index);
        }

        final String header = headerObj.toString();
        if (null != this.headerAlias) {
            return ObjectUtil.defaultIfNull(this.headerAlias.get(header), header);
        }
        return header;
    }

    //    /**
//     * 设置单元格值处理逻辑<br>
//     * 当Excel中的值并不能满足我们的读取要求时，通过传入一个编辑接口，可以对单元格值自定义，例如对数字和日期类型值转换为字符串等
//     *
//     * @param cellEditor 单元格值处理接口
//     */
//    public void setCellEditor(CellEditor cellEditor) {
//        this.mapSheetReader.setCellEditor(cellEditor);
//    }
//
//    /**
//     * 设置是否忽略空行
//     *
//     * @param ignoreEmptyRow 是否忽略空行
//     */
//    public void setIgnoreEmptyRow(boolean ignoreEmptyRow) {
//        this.mapSheetReader.setIgnoreEmptyRow(ignoreEmptyRow);
//    }

    /**
     * 设置标题行的别名Map
     *
     * @param headerAlias 别名Map
     */
    @Override
    public void setHeaderAlias(Map<String, String> headerAlias) {
        if (headerAlias != null) {
            this.headerAlias.putAll(headerAlias);
        }
    }

    /**
     * 增加标题别名
     *
     * @param header 标题
     * @param alias  别名
     */
    @Override
    public void addHeaderAlias(String header, String alias) {
        this.headerAlias.put(header, alias);
    }
}
