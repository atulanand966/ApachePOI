package com.excel.demo;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadDemo {

	public static void main(String[] args) throws IOException {
		
		String filePath = "./target/read_demo.xlsx";
		
		FileInputStream inputStream = new FileInputStream(filePath);
		
		Workbook workbook = new XSSFWorkbook(inputStream);
		
		int totalSheets = workbook.getNumberOfSheets();
		
		System.out.println("Total Number of sheets present is "+totalSheets);
		
		for(int currSheet=0; currSheet<totalSheets; currSheet++) {
			XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(currSheet);
			System.out.println("Current sheet is : "+sheet.getSheetName());
			System.out.println("==========================================");
			int lastRowIndex = sheet.getLastRowNum();
			//System.out.println("lastRowIndex is : "+lastRowIndex);
			for(int rowInd = 0; rowInd <= lastRowIndex; rowInd++) {
				XSSFRow row = sheet.getRow(rowInd);
				int lastCellInd = row.getLastCellNum();
				//System.out.println("lastCellInd is : "+lastCellInd);
				for(int cellInd = 0; cellInd <= lastCellInd; cellInd++) {
					XSSFCell cell = row.getCell(cellInd);
					if(cell != null) {
						switch (cell.getCellType()) {
							case STRING:
								System.out.print(cell.getStringCellValue());
								break;
							case NUMERIC:
								System.out.print(cell.getNumericCellValue());
								break;
							case BOOLEAN:
								System.out.print(cell.getBooleanCellValue());
								break;
							default:
								System.out.println("Unknown");
								break;							
						}
						System.out.print(" | ");
					}
					
				}
				System.out.println();
			}
			System.out.println("==========================================");
		}

	}

}
