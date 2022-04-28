package com.thoughtworks.handlerexcel.config;

import lombok.Data;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

@Data
@Component
@ConfigurationProperties(prefix = "excel.config")
public class ExcelArgumentProperties {

    private String projectCode;

}
