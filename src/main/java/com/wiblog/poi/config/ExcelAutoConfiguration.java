package com.wiblog.poi.config;

import com.wiblog.poi.excel.handler.DefaultExcelDictHandler;
import com.wiblog.poi.excel.handler.IExcelDictHandler;
import com.wiblog.poi.excel.reader.ExcelHandler;
import org.springframework.boot.autoconfigure.condition.ConditionalOnMissingBean;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

/**
 * @author panwm
 * @since 2023/10/21 22:59
 */
@Configuration(proxyBeanMethods = false)
public class ExcelAutoConfiguration {

    @Bean
    public ExcelHandler excelHandler(IExcelDictHandler excelDictHandler) {
        return new ExcelHandler(excelDictHandler);
    }

    @Bean
    @ConditionalOnMissingBean(IExcelDictHandler.class)
    public IExcelDictHandler excelDictHandler() {
        return new DefaultExcelDictHandler();
    }
}
