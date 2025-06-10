package com.lendscape.oa.aiexcel;

import org.springframework.boot.SpringApplication;

public class TestAiExcelApplication {

    public static void main(String[] args) {
        SpringApplication.from(AiExcelApplication::main).with(TestcontainersConfiguration.class).run(args);
    }

}
