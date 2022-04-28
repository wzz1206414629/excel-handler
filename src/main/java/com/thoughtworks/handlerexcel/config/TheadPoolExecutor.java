package com.thoughtworks.handlerexcel.config;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

import java.util.concurrent.Executors;
import java.util.concurrent.LinkedBlockingDeque;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;

@Configuration
public class TheadPoolExecutor {

    @Bean
    public ThreadPoolExecutor threadPoolExecutor() {
        return new ThreadPoolExecutor(4
                , 8
                , 3
                , TimeUnit.SECONDS
                , new LinkedBlockingDeque<>(6)
                , Executors.defaultThreadFactory()
                , new ThreadPoolExecutor.DiscardOldestPolicy());
    }
}
