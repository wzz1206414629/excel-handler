package com.thoughtworks.handlerexcel.model;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class TimeCardModel {

    @ExcelProperty(index = 1)
    private String name;

    @ExcelProperty(index = 3)
    private String project;

    @ExcelProperty(index = 4)
    private String subproject;

    @ExcelProperty(index = 5)
    private Date date;

    @ExcelProperty(index = 6)
    private double billableHour;

    @ExcelProperty(index = 7)
    private double nonBillableHour;

}
