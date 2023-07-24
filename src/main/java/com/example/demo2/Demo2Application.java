package com.example.demo2;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class Demo2Application {

	public static void main(String[] args) {
		SpringApplication.run(Demo2Application.class, args);
		try {
			ExcelGenerator.generateExcelFile();
			System.out.println("Excel file generated successfully.");
		} catch (Exception e) {
			System.out.println("Error generating Excel file: " + e.getMessage());
		}
	}

}
