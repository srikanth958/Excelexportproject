package com.app.export;

import java.io.File;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.inno.exports.ExcelExport;
import com.inno.exports.SuperExport;

public class DbToExcel {

	public  static void main(String[] args)throws Exception {
		Connection con =DriverManager.getConnection("jdbc:mysql://localhost:3306/test?user=root&password=xxxx");
	Statement stmt=con.createStatement();
	stmt.executeQuery("select *from test.csvs");
	ResultSet rs=null;
	FileOutputStream fout=new FileOutputStream("C:\\Users\\Lenova\\Desktop\\db.xlsx");
	XSSFWorkbook workbook =new XSSFWorkbook();
	XSSFSheet sheet=workbook.createSheet("dbData");
	Row row=sheet.createRow(0);
	Cell cell=row.createCell(0);
	cell.setCellValue("name");
	Cell cell1=row.createCell(1);
	cell1.setCellValue("ptype");
	Cell cell2=row.createCell(2);
	cell2.setCellValue("prange");
	Cell cell3=row.createCell(3);
	cell3.setCellValue("quantity");
	Cell cell4=row.createCell(4);
	cell4.setCellValue("avalil");
	populte(rs, workbook, sheet);
	workbook.write(fout);
	fout.close();
	con.close();
		

	}
	public static void populte(ResultSet rs,XSSFWorkbook workbook,XSSFSheet sheet) throws SQLException {
		int rowcount=0;
		while(rs.next()) {
			String v1=rs.getString(1);
			String v2=rs.getString(2);
			String v3=rs.getString(3);
			String v4=rs.getString(4);
			String v5=rs.getString(5);
			
			
			Row row=sheet.createRow(++rowcount);
			 int columncount=0;
			 Cell cell1=row.createCell(columncount++);
			 cell1.setCellValue(v1);
			 Cell cell2=row.createCell(columncount++);
			 cell2.setCellValue(v2);
			 Cell cell3=row.createCell(columncount++);
			 cell3.setCellValue(v3);
			 Cell cell4=row.createCell(columncount++);
			 cell4.setCellValue(v4);
			 Cell cell5=row.createCell(columncount++);
			 cell5.setCellValue(v5);
		}
		
	}

}

