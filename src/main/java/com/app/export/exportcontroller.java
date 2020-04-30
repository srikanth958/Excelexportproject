package com.app.export;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class exportcontroller  {
	public  void main(String args[]) throws Exception {
	FileOutputStream fos=new FileOutputStream(new File("C:\\Users\\Lenova\\Desktop\\sample.xlsx"));
	 XSSFWorkbook workbook = new XSSFWorkbook();
     XSSFSheet sheet = workbook.createSheet("Java Books");
      
     Object[][] bookData = {
             {"Head First Java", "Kathy Serria", 79},
             {"Effective Java", "Joshua Bloch", 36},
             {"Clean Code", "Robert martin", 42},
             {"Thinking in Java", "Bruce Eckel", 35},
     };

     int rowCount = 0;
      
     for (Object[] aBook : bookData) {
         Row row = sheet.createRow(++rowCount);
          
         int columnCount = 0;
          
         for (Object field : aBook) {
             Cell cell = row.createCell(++columnCount);
             if (field instanceof String) {
                 cell.setCellValue((String) field);
             } else if (field instanceof Integer) {
                 cell.setCellValue((Integer) field);
             }
         }
          
     }
     workbook.write(fos);
	}
}
