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
		Connection con =DriverManager.getConnection("jdbc:mysql://localhost:3306/test?user=root&password=9580");
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
	//populte(rs, workbook, sheet);
	beans bean=new beans("ram", 23, 75.87d, 87.8676f, true);
	beans bean1=new beans("ram", 23, 75.87d, 87.8676f, true);
	beans bean2=new beans();
	beans bean3=new beans();
	List<beans> bns= Arrays.asList(bean2,bean3);
	ArrayList arr=new ArrayList();
	@SuppressWarnings("rawtypes")
	SuperExport excelExport=new ExcelExport<beans>(); 
	//excelExport.Export(bean2, "check2.xlsx");
	 //excelExport.Export(rs, "sample2.xlsx");
	excelExport.Export(arr, "check3.xlsx");
	
    
	
	//workbook.write(fout);
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
class beans{
	
	private String name;
	private int age;
	private double Double;
	private float ft;
	private boolean bool;
	public beans() {
		
	}
	public beans(String name, int age, double d, float ft, boolean bool) {
		super();
		this.name = name;
		this.age = age;
		Double = d;
		this.ft = ft;
		this.bool = bool;
	}
	
	
}
