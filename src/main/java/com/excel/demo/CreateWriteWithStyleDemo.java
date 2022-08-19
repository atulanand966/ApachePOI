package com.excel.demo;


import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateWriteWithStyleDemo {

	public static void main(String[] args) throws IOException {
		
		String filePath = "./target/createWriteWithStyleDemo.xlsx";
		
		Object[][] empDetailsArr = {
				{"Emp Id","Name","Salary(INR)"},
				{1001,"Ram Kumar",60000},
				{1002,"John Wick",55000},
				{1003,"Thomas Edition",73000},
				{1004,"Jason Bourne",97000}
		};
		
		Object[][] studentDetailsArr = {
				{"Student Id","Name","Buddy"},
				{"UR12CS019","Ankit Kumar","Priyadharsini C"},
				{"UR12CS022","Anshul Tripathi","Priyadharsini C"},
				{"UR12CS043","Cincy Sebastian","Sebastian Terrence"},
				{"UR12CS044","Clinton Thomas","Sebastian Terrence"},
		};
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet empDetailsSheet = workbook.createSheet("Employee Details");
		XSSFSheet studentDetailsSheet = workbook.createSheet("Student Details");
		
		XSSFCellStyle headerStyle = workbook.createCellStyle();
		headerStyle.setAlignment(HorizontalAlignment.CENTER);
		headerStyle.setBorderTop(BorderStyle.MEDIUM);
		headerStyle.setBorderLeft(BorderStyle.MEDIUM);
		headerStyle.setBorderRight(BorderStyle.MEDIUM);
		headerStyle.setBorderBottom(BorderStyle.MEDIUM);
		headerStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
		headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		XSSFFont headerFont = workbook.createFont();
		headerFont.setBold(true);
		headerFont.setFontName("Arial");
		headerStyle.setFont(headerFont);
		
		
		XSSFCellStyle dataStyle = workbook.createCellStyle();
		dataStyle.setAlignment(HorizontalAlignment.LEFT);
		dataStyle.setBorderTop(BorderStyle.THIN);
		dataStyle.setBorderLeft(BorderStyle.THIN);
		dataStyle.setBorderRight(BorderStyle.THIN);
		dataStyle.setBorderBottom(BorderStyle.THIN);
		
		
		int rowNum = 0;
		for(Object[] emp : empDetailsArr) {
			XSSFRow row = empDetailsSheet.createRow(rowNum);
			int colNum = 0;
			for(Object empData : emp) {
				XSSFCell cell = row.createCell(colNum++);
				
				if(rowNum == 0) {
					cell.setCellStyle(headerStyle);
				}
				else {
					cell.setCellStyle(dataStyle);
				}
				
				if(empData instanceof String) {
					cell.setCellValue((String) empData);
				}
				else if(empData instanceof Integer) {
					cell.setCellValue((Integer) empData);
				}
			}
			rowNum++;
		}
		
		
		rowNum = 0;
		for(Object[] emp : studentDetailsArr) {
			XSSFRow row = studentDetailsSheet.createRow(rowNum);
			int colNum = 0;
			for(Object empData : emp) {
				XSSFCell cell = row.createCell(colNum++);
				
				if(rowNum == 0) {
					cell.setCellStyle(headerStyle);
				}
				else {
					cell.setCellStyle(dataStyle);
				}
				
				if(empData instanceof String) {
					cell.setCellValue((String) empData);
				}
				else if(empData instanceof Integer) {
					cell.setCellValue((Integer) empData);
				}
			}
			rowNum++;
		}
		
		FileOutputStream outputStream = new FileOutputStream(filePath);
		workbook.write(outputStream);
		outputStream.close();
		workbook.close();
		System.out.println("Excel File Created Successfully.....!!!");
	}

}
