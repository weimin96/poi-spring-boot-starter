package com.wiblog.poi;

import com.wiblog.poi.bean.MergeScore;
import com.wiblog.poi.bean.Score;
import com.wiblog.poi.excel.ExportParam;
import com.wiblog.poi.excel.reader.ExcelHandler;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.util.ResourceUtils;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.List;

@SpringBootTest
class PoiSpringBootStarterApplicationTests {

    @Autowired
    private ExcelHandler excelHandler;

    public static final String EXCEL_PATH = "data/test1.xlsx";

    public static final String MERGE_EXCEL_PATH = "data/mergetest.xlsx";

    @Test
    void readExcel() throws FileNotFoundException {
        File file = ResourceUtils.getFile(ResourceUtils.CLASSPATH_URL_PREFIX + EXCEL_PATH);
        List<Score> scores = excelHandler.readExcel(file, Score.class);
        System.out.println();
    }

    @Test
    void readMergeExcel() throws FileNotFoundException {
        File file = ResourceUtils.getFile(ResourceUtils.CLASSPATH_URL_PREFIX + MERGE_EXCEL_PATH);
        List<MergeScore> scores = excelHandler.readExcel(file, 0, 1, 4, MergeScore.class);
        System.out.println();
    }

    @Test
    void export() throws IOException {
        File file = ResourceUtils.getFile(ResourceUtils.CLASSPATH_URL_PREFIX + EXCEL_PATH);
        List<Score> scores = excelHandler.readExcel(file, Score.class);
        String path = file.getParentFile().getAbsolutePath() + File.separator + "export.xlsx";

        File expportFile = new File(path);
        expportFile.delete();

        ExportParam param = ExportParam.builder().data(scores).entity(Score.class).build();
        excelHandler.writeExcel(param, expportFile);
    }

}
