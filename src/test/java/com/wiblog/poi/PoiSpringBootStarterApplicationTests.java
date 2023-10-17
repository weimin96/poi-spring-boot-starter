package com.wiblog.poi;

import cn.hutool.core.date.DateUtil;
import cn.hutool.core.io.FileUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import com.wiblog.poi.bean.Score;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.util.ResourceUtils;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;

@SpringBootTest
class PoiSpringBootStarterApplicationTests {

    public static final String EXCEL_PATH = "data/test1.xlsx";

    @Test
    void readExcel() throws FileNotFoundException {
        File file = ResourceUtils.getFile(ResourceUtils.CLASSPATH_URL_PREFIX + EXCEL_PATH);
        ExcelReader reader = ExcelUtil.getReader(FileUtil.file(file));
        reader.addHeaderAlias("姓名", "name");
        reader.addHeaderAlias("班级", "className");
        reader.addHeaderAlias("科目", "subject");
        reader.addHeaderAlias("成绩", "score");
        List<Score> scores = reader.readAll(Score.class);
        System.out.println(scores);
    }

    @Test
    void export() throws IOException {
        File file = ResourceUtils.getFile(ResourceUtils.CLASSPATH_URL_PREFIX + EXCEL_PATH);
        ExcelReader reader = ExcelUtil.getReader(FileUtil.file(file));
        reader.addHeaderAlias("姓名", "name");
        reader.addHeaderAlias("班级", "className");
        reader.addHeaderAlias("科目", "subject");
        reader.addHeaderAlias("成绩", "score");
        List<Score> scores = reader.readAll(Score.class);
        String path = file.getParentFile().getAbsolutePath() + File.separator + "testExport.xlsx";

        File expportFile = new File(path);
        expportFile.delete();
        // 通过工具类创建writer
        ExcelWriter writer = ExcelUtil.getWriter(expportFile);
        // 合并单元格后的标题行，使用默认标题样式
        writer.merge(4, "成绩单");
        writer.setColumnWidth(1, 30);

        // 一次性写出内容，使用默认样式，强制输出标题
        writer.write(scores, true);
        // 关闭writer，释放内存
        writer.close();

    }

}
