package com.excel.demo;


import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
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

public class DataFormatDemo {

	public static void main(String[] args) throws IOException, ParseException {
		
		String filePath = "./target/dataFormatDemo.xlsx";
		
		SimpleDateFormat dateFormat = new SimpleDateFormat("dd MMM, yyyy");
		
		Object[][] empDetailsArr = {
				{"Emp Id","Name","Salary(INR)", "D.O.B."},
				{1001,"Ram Kumar",60000,dateFormat.parse("23 Oct, 1991")},
				{1002,"John Wick",55000,dateFormat.parse("07 Jan, 1987")},
				{1003,"Thomas Edition",73000,dateFormat.parse("14 Nov, 1982")},
				{1004,"Jason Bourne",97000,dateFormat.parse("11 Dec, 1993")}
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
		
		XSSFCellStyle dateStyle = workbook.createCellStyle();
		dateStyle.setAlignment(HorizontalAlignment.RIGHT);
		dateStyle.setBorderTop(BorderStyle.THIN);
		dateStyle.setBorderLeft(BorderStyle.THIN);
		dateStyle.setBorderRight(BorderStyle.THIN);
		dateStyle.setBorderBottom(BorderStyle.THIN);
		CreationHelper creationHelper= workbook.getCreationHelper();
		dateStyle.setDataFormat(creationHelper.createDataFormat().getFormat("dd-MM-yyyy"));
		
		
		int rowNum = 0;
		for(Object[] emp : empDetailsArr) {
			XSSFRow row = empDetailsSheet.createRow(rowNum);
			int colNum = 0;
			for(Object empData : emp) {
				XSSFCell cell = row.createCell(colNum++);
				
								
				
				if(empData instanceof String) {
					if(rowNum == 0) {
						cell.setCellStyle(headerStyle);
					}else {
						cell.setCellStyle(dataStyle);
					}
					
					cell.setCellValue((String) empData);
				}
				else if(empData instanceof Integer) {
					cell.setCellStyle(dataStyle);
					cell.setCellValue((Integer) empData);
				}
				else if(empData instanceof Date) {
					cell.setCellStyle(dateStyle);
					cell.setCellValue((Date) empData);
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
