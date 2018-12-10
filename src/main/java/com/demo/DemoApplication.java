package com.demo;

import com.biz.GetPageSourceBiz;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.autoconfigure.jdbc.DataSourceAutoConfiguration;

@SpringBootApplication(exclude = {DataSourceAutoConfiguration.class})
public class DemoApplication {

    @Autowired
    private static GetPageSourceBiz getPageSourceBiz;

    public static void main(String[] args) {


        SpringApplication.run(DemoApplication.class, args);
        getPageSourceBiz.pageSource();

    }
}
