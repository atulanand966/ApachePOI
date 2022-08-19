package com.excel.demo;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadIteratorDemo {

	public static void main(String[] args) throws IOException {
		
		String filePath = "./target/read_demo.xlsx";
		
		FileInputStream inputStream = new FileInputStream(filePath);
		
		Workbook workbook = new XSSFWorkbook(inputStream);
		
		Iterator<Sheet> sheetIterator = workbook.sheetIterator();
		while(sheetIterator.hasNext()) {
			XSSFSheet sheet = (XSSFSheet) sheetIterator.next();
			System.out.println("Current sheet is : "+sheet.getSheetName());
			System.out.println("********************************************");
			Iterator<Row> rowIterator = sheet.rowIterator();
			while(rowIterator.hasNext()) {
				XSSFRow row = (XSSFRow) rowIterator.next();
				Iterator<Cell> cellIterator = row.cellIterator();
				while(cellIterator.hasNext()) {
					XSSFCell cell = (XSSFCell) cellIterator.next();
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
					System.out.print(" || ");
				}
				System.out.println();
			}
			System.out.println("********************************************");
		}

	}

}
