package com.task.xlsx_task;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Insert {

	public static void main(String[] args) throws Exception{
		XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream("C:\\Users\\DICKSON\\Desktop\\Order_Summary_Template.xlsx"));
        XSSFSheet sheet = workbook.getSheetAt(0);
        
        //Getting the number of rows
        int lastRow = sheet.getLastRowNum();
        
        //I created an empty  row between 4 and 5 
        sheet.shiftRows(3, lastRow, 1);
        sheet.createRow(3);
        
        //I copy row 9 into the shifted row
        sheet.copyRows(9, 9, 3, new CellCopyPolicy());
       
        
		
        FileOutputStream out = new FileOutputStream("C:\\Users\\DICKSON\\Desktop\\Final_Order_Summary_Template.xlsx");
        workbook.write(out);
        out.close();
    
	
}
}
