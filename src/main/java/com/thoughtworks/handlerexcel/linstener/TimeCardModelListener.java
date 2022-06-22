package com.thoughtworks.handlerexcel.linstener;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.read.listener.ReadListener;
import com.alibaba.excel.util.BooleanUtils;
import com.alibaba.excel.util.ListUtils;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.handler.context.CellWriteHandlerContext;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.style.column.SimpleColumnWidthStyleStrategy;
import com.thoughtworks.handlerexcel.config.ExcelArgumentProperties;
import com.thoughtworks.handlerexcel.model.TimeCardModel;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.springframework.stereotype.Component;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

@Slf4j
@Component
@RequiredArgsConstructor
public class TimeCardModelListener implements ReadListener<TimeCardModel> {

    private final ExcelArgumentProperties properties;

    private final ThreadPoolExecutor threadPoolExecutor;

    public HttpServletResponse response;

    public void setResponse(HttpServletResponse response) {
        this.response = response;
    }

    /**
     * 单次缓存的数据量
     */
    public static final int BATCH_COUNT = 100;

    /**
     * 临时存储
     */
    private List<TimeCardModel> cachedDataList = ListUtils.newArrayListWithExpectedSize(BATCH_COUNT);

    @Override
    public void invoke(TimeCardModel timeCardModel, AnalysisContext analysisContext) {
        cachedDataList.add(timeCardModel);
        if (cachedDataList.size() >= BATCH_COUNT) {
            // 存储完成 筛选覆盖list
            cachedDataList = cachedDataList.stream()
                    .filter(model -> properties.getProjectCode().equals(model.getProject()))
                    .collect(Collectors.toList());
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        try {
            export();
        } catch (Exception ignore) {
        }
    }


    public void export() throws ExecutionException, InterruptedException, IOException {
        Date date = cachedDataList.get(0).getDate();
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(date);
        int monthDay = calendar.getActualMaximum(Calendar.DAY_OF_MONTH);

        // 获取excel自定义的表格头信息 (2022/4/1, 2022/4/2)
        List<List<String>> monthDayExcelHeader = new ArrayList<>();
        CompletableFuture<Void> excelHeaderFuture = CompletableFuture.runAsync(() -> monthDayExcelHeader.addAll(getExcelHeader(calendar)),
                threadPoolExecutor);

        // 自定义excel的表格体信息
        List<List<Object>> statisticsData = new ArrayList<>();
        CompletableFuture<Void> excelBodyFuture = CompletableFuture.runAsync(() -> statisticsData.addAll(getExcelBody(calendar)),
                threadPoolExecutor);

        // 等待异步任务返回header 和 body
        CompletableFuture.allOf(excelBodyFuture, excelHeaderFuture).get();

        // 自定义样式
        CellWriteHandler cellWriteHandler = diyCellStyle(monthDay);

        EasyExcel.write(response.getOutputStream())
                .head(monthDayExcelHeader)
                .sheet()
                // 列宽 20
                .registerWriteHandler(new SimpleColumnWidthStyleStrategy(20))
                // 自定义
                .registerWriteHandler(cellWriteHandler)
                .doWrite(statisticsData);
    }

    private void setColor(WriteCellStyle writeCellStyle, IndexedColors colors) {
        writeCellStyle.setFillForegroundColor(colors.getIndex());
        // 这里需要指定 FillPatternType 为FillPatternType.SOLID_FOREGROUND
        writeCellStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
    }

    private CellWriteHandler diyCellStyle(int monthDay) {
        List<Integer> weekendColumns = new ArrayList<>();
        return new CellWriteHandler() {
            @Override
            public void afterCellDispose(CellWriteHandlerContext context) {
                // 当前事件会在 数据设置到poi的cell里面才会回调
                // 判断不是头的情况 如果是fill 的情况 这里会==null 所以用not true
                if (BooleanUtils.isNotTrue(context.getHead())) {
                    // 第一个单元格
                    // 只要不是头 一定会有数据 当然fill的情况 可能要context.getCellDataList() ,这个需要看模板，因为一个单元格会有多个 WriteCellData
                    WriteCellData<?> cellData = context.getFirstCellData();
                    if (cellData.getColumnIndex() >= 1 && cellData.getColumnIndex() <= monthDay) {
                        double cellValue = cellData.getNumberValue().doubleValue();
                        if (cellData.getRowIndex() == 1 && cellValue == 0) {
                            // 当前是「应当工作时间的列」& 当前是周末
                            setColor(cellData.getOrCreateStyle(), IndexedColors.GREY_25_PERCENT);
                            // 周末的列存储 供给 每个人的周末高亮使用
                            weekendColumns.add(cellData.getColumnIndex());
                        } else {// 否则是每个人的工作时间
                            if (weekendColumns.contains(cellData.getColumnIndex())) {
                                // 当前是周末
                                setColor(cellData.getOrCreateStyle(), IndexedColors.GREY_25_PERCENT);
                            } else {
                                // 不是周末的情况
                                if (cellValue < 8) {
                                    setColor(cellData.getOrCreateStyle(), IndexedColors.RED);
                                } else if (cellValue > 8) {
                                    setColor(cellData.getOrCreateStyle(), IndexedColors.GREEN);
                                }
                            }
                        }
                    }
                }
            }
        };
    }

    private List<List<Object>> getExcelBody(Calendar calendar) {
        int monthDay = calendar.getActualMaximum(Calendar.DAY_OF_MONTH);
        List<List<Object>> statisticsData = new ArrayList<>();
        // 第一行为 「应当工作时间 8 8 8 ... 0 0 8 8 8」
        List<Object> firstRow = new ArrayList<>();
        firstRow.add("应当工作时间");
        AtomicInteger totalShouldWorkHour = new AtomicInteger();
        IntStream.range(1, monthDay + 1)
                .forEach(day -> {
                    calendar.set(Calendar.DAY_OF_MONTH, day);
                    int shouldWorkHour = calendar.get(Calendar.DAY_OF_WEEK) == Calendar.SATURDAY
                            || calendar.get(Calendar.DAY_OF_WEEK) == Calendar.SUNDAY ? 0 : 8;
                    totalShouldWorkHour.addAndGet(shouldWorkHour);
                    firstRow.add(shouldWorkHour);
                });
        firstRow.add(totalShouldWorkHour.get());
        firstRow.add(((double) (totalShouldWorkHour.get())) / 8);
        statisticsData.add(firstRow);

        // 根据每个人的姓名分组
        Map<String, List<TimeCardModel>> usernameMap = cachedDataList.stream()
                .collect(Collectors.groupingBy(TimeCardModel::getName));
        usernameMap.forEach((username, timeCardModels) -> {
            // 记录每个人每天的工作时间 eg: zhouzhou wang 8 8 8 ... 0 0 160 20
            List<Object> statisticsDataItem = new ArrayList<>(31);
            statisticsDataItem.add(username);
            // 月的每天工作时长
            double[] monthWorkHour = new double[monthDay];
            for (TimeCardModel timeCardModel : timeCardModels) {
                Calendar currentTime = Calendar.getInstance();
                currentTime.setTime(timeCardModel.getDate());
                // 记录工作的时间
                double workTime = timeCardModel.getBillableHour() + timeCardModel.getNonBillableHour();
                monthWorkHour[currentTime.get(Calendar.DAY_OF_MONTH) - 1] = workTime;
            }
            // 把每个人每天的工作时间放入list中
            statisticsDataItem.addAll(Arrays.stream(monthWorkHour)
                    .boxed()
                    .collect(Collectors.toList()));
            // 放入总共工作时间
            double totalWorkHour = Arrays.stream(monthWorkHour).sum();
            statisticsDataItem.add(totalWorkHour);
            // 放入工作人天
            statisticsDataItem.add(totalWorkHour / 8);
            statisticsData.add(statisticsDataItem);
        });
        return statisticsData;
    }

    private List<List<String>> getExcelHeader(Calendar calendar) {
        int monthDay = calendar.getActualMaximum(Calendar.DAY_OF_MONTH);
        List<List<String>> monthDayExcelHeader = new ArrayList<>(31);
        monthDayExcelHeader.add(Collections.singletonList(""));
        IntStream.range(1, monthDay + 1)
                .forEach(dayOfMonth -> monthDayExcelHeader.add(
                        Collections.singletonList(
                                LocalDate.of(calendar.get(Calendar.YEAR), calendar.get(Calendar.MONTH) + 1, dayOfMonth)
                                        .format(DateTimeFormatter.ofPattern("yyyy/M/d"))
                        )
                ));
        monthDayExcelHeader.add(Collections.singletonList("总共工作小时"));
        monthDayExcelHeader.add(Collections.singletonList("总共工作人天"));
        System.out.println(monthDayExcelHeader);
        return monthDayExcelHeader;
    }
}

