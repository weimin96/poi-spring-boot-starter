package com.wiblog.poi.excel.reader;

import ch.qos.logback.classic.util.LogbackMDCAdapter;
import cn.hutool.core.bean.BeanUtil;
import cn.hutool.core.bean.copier.CopyOptions;
import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.collection.IterUtil;
import cn.hutool.core.collection.ListUtil;
import cn.hutool.core.date.DateUtil;
import cn.hutool.core.util.ObjectUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.RowUtil;
import cn.hutool.poi.excel.cell.CellUtil;
import cn.hutool.poi.excel.reader.AbstractSheetReader;
import com.wiblog.poi.excel.annotation.Excel;
import com.wiblog.poi.excel.bean.MergeCell;
import com.wiblog.poi.excel.handler.IExcelDictHandler;
import com.wiblog.poi.exception.ExcelErrorException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.StringUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;

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

    private final IExcelDictHandler dictHandler;

    private final Map<String, String> headerAlias;

    private final Map<String, MergeCell> mergeCellMap;

    private final Map<String, String> dictTypeMap;

    private final Map<String, Map<String, String>> replaceMap;

    private final Map<String, String> dateFormatMap;

    private final Map<String, Map<String, Object>> numFormatMap;

    private final Map<Integer, String> orderMap = new HashMap<>();

    /**
     * 合并单元格的列集合
     */
    private final Set<Integer> mergeSet = new HashSet<>();

    private Sheet sheet;

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
        this.sheet = sheet;
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
                setExcelMap(field);
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
        // 索引-》field
        for (int i = 0; i < headerList.size(); i++) {
            this.orderMap.put(i, headerList.get(i));
        }

        for (Field field : fields) {
            if (field.isAnnotationPresent(Excel.class)) {
                Excel annotation = field.getAnnotation(Excel.class);
                if (annotation.merge()) {
                    MergeCell mergeCell = new MergeCell();
                    this.mergeCellMap.put(field.getName(), mergeCell);
                    setMergeCell(field, mergeCell, null, 1);
                }
            }
        }

        final List<T> beanList = new ArrayList<>(endRowIndex - startRowIndex + 1);
        final CopyOptions copyOptions = CopyOptions.create().setIgnoreError(true);
        List<Object> rowList;
        for (int i = startRowIndex; i <= endRowIndex; i++) {
            // 跳过标题行
            rowList = readRow(sheet, i);
            if (CollUtil.isNotEmpty(rowList) || !ignoreEmptyRow) {

                // 值替换
                    handlerReplace(rowList);
                // 设置翻译
                handlerTranslators(rowList);
                // 设置日期格式
                    handlerDateFormat(rowList);
                // 设置数值格式化
                    handlerNumFormat(rowList);
                Map<String, Object> dataMap = initDataMap(rowList);
//                Map<String, Object> dataMap = IterUtil.toMap(headerList, rowList, true);
                // 合并单元格处理
//                for (String mergeName : mergeList) {
//                    initializeFields(this.mergeCellMap.get(mergeName), dataMap, 0, new HashMap<>());
//                }

                T bean = BeanUtil.toBean(dataMap, this.beanClass, copyOptions);

                beanList.add(bean);
            }
        }
        return beanList;
    }

    /**
     * list转map
     * @param rowList rowList
     * @return Map
     */
    public Map<String, Object> initDataMap(List<Object> rowList) {
        Map<String, Object> dataMap = new HashMap<>(32);
        for (int i = 0; i < rowList.size(); i++) {
            String fieldName = this.orderMap.get(i);
            Object o = rowList.get(i);
            // 该列不是合并单元格列
            if (!mergeSet.contains(i)) {
                dataMap.put(fieldName, o);
            }
        }
        // 处理合并单元格
        for (String fieldName: this.mergeCellMap.keySet()) {
            Map<String,Object> map = new HashMap<>();
            MergeCell mergeCell = this.mergeCellMap.get(fieldName);
            buildMergeMap(map, mergeCell, rowList);
            dataMap.putAll(map);
        }
        return dataMap;
    }

    public void buildMergeMap(Map<String,Object> map, MergeCell mergeCell, List<Object> rowList) {
        Map<String, MergeCell> children = mergeCell.getChildren();
        if (children == null) {
            int index = mergeCell.getIndex();
            Object o = rowList.get(index);
            map.put(mergeCell.getField(), o);
            return;
        }
        Map<String,Object> childrenMap = new HashMap<>(32);
        map.put(mergeCell.getField(), childrenMap);
        for (String fieldName: children.keySet()) {
            MergeCell child = children.get(fieldName);
            buildMergeMap(childrenMap, child, rowList);
        }
    }

    /**
     * 合并单元格映射处理
     *
     * @param field field
     */
    public void setMergeCell(Field field, MergeCell mergeCell, MergeCell parentCell, int level) {
        Excel annotation = field.getAnnotation(Excel.class);
        mergeCell.setField(field.getName());
        mergeCell.setName(annotation.name());
        if (!annotation.merge()) {
            boolean repect = false;
            int index = -1;
            List<Object> objects = readRow(this.sheet, headerRowIndex + level - 2);
            for (int i = 0; i < this.orderMap.size(); i++) {
                if (this.orderMap.get(i).equals(field.getName())) {
                    if (index != -1) {
                        repect = true;
                    }
                    // 多个同名字段
                    if (repect) {
                        String parentName = parentCell.getName();

                        String parentHeaderName = (String) objects.get(i);
                        if (parentName.equals(parentHeaderName)) {
                            index = i;
                            break;
                        }
                    } else {
                        index = i;
                    }
                }
            }
            this.mergeSet.add(index);
            mergeCell.setIndex(index);
            return;
        }

        Map<String, MergeCell> children = new HashMap<>(32);
        mergeCell.setChildren(children);

        Field[] fields = field.getType().getDeclaredFields();
        for (Field child : fields) {
            if (!child.isAnnotationPresent(Excel.class)) {
                continue;
            }
            MergeCell childCell = new MergeCell();
            setMergeCell(child, childCell, mergeCell, level + 1);
            children.put(child.getName(), childCell);
        }
    }

    public void handlerReplace(List<Object> row) {
        for (int i = 0; i < row.size(); i++) {
            String key = this.orderMap.get(i);
            Map<String, String> repMap = replaceMap.get(key);
            if (repMap != null) {
                String val = ObjectUtil.toString(row.get(i));
                for (String mapKey : repMap.keySet()) {
                    if (mapKey.equals(val)) {
                        row.set(i, repMap.get(mapKey));
                    }
                }
            }
        }
    }

    public void handlerTranslators(List<Object> row) {
        for (int i = 0; i < row.size(); i++) {
            String dictType = dictTypeMap.get(this.orderMap.get(i));
            if (StringUtil.isNotBlank(dictType)) {
                Object o = row.get(i);
                if (o != null) {
                    row.set(i, dictHandler.toValue(dictType, ObjectUtil.toString(o)));
                }
            }
        }
    }

    public void handlerDateFormat(List<Object> row) {
        for (int i = 0; i < row.size(); i++) {
            String filedName = this.orderMap.get(i);
            String dataFormat = dateFormatMap.get(filedName);
            if (StringUtil.isNotBlank(dataFormat)) {
                try {
                    Date date = (Date) row.get(i);
                    row.set(i, DateUtil.format(date, dataFormat));
                } catch (Exception ignored) {
                }
            }
        }
    }

    public void handlerNumFormat(List<Object> row) {
        for (int i = 0; i < row.size(); i++) {
            String filedName = this.orderMap.get(i);
            Map<String, Object> formatMap = numFormatMap.get(filedName);

            if (formatMap != null) {
                String numFormat = (String) formatMap.get("numFormat");
                RoundingMode round = (RoundingMode) formatMap.get("round");
                Object o = row.get(i);
                if (o == null) {
                    continue;
                }
                DecimalFormat df = new DecimalFormat(numFormat);
                df.setRoundingMode(round);
                row.set(i, df.format(o));
            }
        }
    }

    public void setDictTypeMap(Field field, Excel annotation) {
        String dictType = annotation.dictType();
        if (StringUtil.isNotBlank(dictType)) {
            dictTypeMap.put(field.getName(), dictType);
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

    public void setExcelMap(Field field) {
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

        Field[] fields = field.getType().getDeclaredFields();
        for (Field child : fields) {
            setExcelMap(child);
        }
    }


    /*public void initializeFields(MergeCell mergeCell, Map<String, Object> dataMap, int index, Map<String, Object> fieldMap) {
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
        }*/


//    }

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
