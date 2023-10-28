package com.wiblog.poi.util;

import cn.hutool.core.exceptions.DependencyException;
import cn.hutool.core.util.ObjectUtil;
import cn.hutool.poi.PoiChecker;
import cn.hutool.poi.excel.ExcelReader;
import com.wiblog.poi.excel.reader.PoiExcelReader;

import java.io.File;
import java.io.InputStream;

/**
 * @author panwm
 * @since 2023/10/22 15:02
 */
public class PoiExcelUtil {

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
}
