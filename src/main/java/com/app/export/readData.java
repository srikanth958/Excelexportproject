package com.app.export;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class readData {

	public  void main(String[] args) throws Exception {
		FileInputStream fin=new FileInputStream(new File("C:\\Users\\Lenova\\Desktop\\bean.xlsx"));
XSSFWorkbook workbook=new XSSFWorkbook(fin);
XSSFSheet sheet=workbook.getSheetAt(0);
Iterator<Row> rowIterator = sheet.iterator();
while (rowIterator.hasNext()) 
{
    Row row = rowIterator.next();
    //For each row, iterate through all the columns
    Iterator<Cell> cellIterator = row.cellIterator();
     
    while (cellIterator.hasNext()) 
    {
        Cell cell = cellIterator.next();
        //Check the cell type and format accordingly
       
        if (cell.getCellType()==CellType.NUMERIC) {
        
                System.out.print(cell.getNumericCellValue() + "    ");
        }
        else  if (cell.getCellType()==CellType.STRING) {
                System.out.print(cell.getStringCellValue() + "    ");
                
        }
    }
    System.out.println("");
}

		
	}

}
