package com.task.xlsx_task;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Read {

	public static void main(String[] args) throws IOException {

		
		try {
			File file = new File("C:\\Users\\DICKSON\\Desktop\\Order_Summary_Template.xlsx");
			FileInputStream fileInputStream = new FileInputStream(file);
			XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
			XSSFSheet sheet = workbook.getSheetAt(0);
			
			Iterator<Row> iterator = sheet.iterator(); 
			
			while(iterator.hasNext()) {
				Row row = iterator.next();
				
				Iterator<Cell> iterator2 = row.cellIterator();
				
				while(iterator2.hasNext()) {
					Cell cell = iterator2.next();
					
					
					switch(cell.getCellType()) {
						case STRING:
							System.out.println(cell.getStringCellValue() + "\t\t");
							break;
						case NUMERIC:
							System.out.println(cell.getNumericCellValue() + "\t\t");
							break;
						case BOOLEAN:
							System.out.println(cell.getBooleanCellValue() + "\t\t");
							break;
						case BLANK:
							break;
						default:
							
					}
				}
				System.out.println("");
			}
			
		
			fileInputStream.close();
			FileOutputStream fileOutputStream = new FileOutputStream(file);
			workbook.write(fileOutputStream);
			fileOutputStream.close();
			workbook.close();
			
			
		} catch (FileNotFoundException e) {
			
			
		e.printStackTrace();
		}
	
		
	}

}
