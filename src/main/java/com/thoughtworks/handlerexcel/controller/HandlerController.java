package com.thoughtworks.handlerexcel.controller;

import com.alibaba.excel.EasyExcel;
import com.thoughtworks.handlerexcel.linstener.TimeCardModelListener;
import com.thoughtworks.handlerexcel.model.TimeCardModel;
import lombok.RequiredArgsConstructor;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.net.URLEncoder;

@RestController
@RequiredArgsConstructor
public class HandlerController {

    private final TimeCardModelListener timeCardModelListener;

    @PostMapping("upload")
    public void uploadAndDownload(MultipartFile file,
                         HttpServletResponse response) throws IOException {
        // 这里注意 有同学反应使用swagger 会导致各种问题，请直接用浏览器或者用postman
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setCharacterEncoding("utf-8");
        // 这里URLEncoder.encode可以防止中文乱码 当然和easyexcel没有关系
        String fileName = URLEncoder.encode("测试", "UTF-8").replaceAll("\\+", "%20");
        response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");

        timeCardModelListener.response = response;
        EasyExcel.read(file.getInputStream(), TimeCardModel.class, timeCardModelListener)
                .headRowNumber(2)
                .sheet()
                .doRead();
    }
}
