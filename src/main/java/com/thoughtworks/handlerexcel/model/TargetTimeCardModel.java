package com.thoughtworks.handlerexcel.model;

import com.alibaba.excel.write.handler.CellWriteHandler;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

@Data
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class TargetTimeCardModel {

    // 自定义表头
    private List<List<String>> monthDayExcelHeader;

    // 自定义样式
    private CellWriteHandler cellWriteHandler;

    // 表格体
    private List<List<Object>> statisticsData;
}
