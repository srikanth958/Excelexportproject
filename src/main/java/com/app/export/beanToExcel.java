package com.app.export;


import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class beanToExcel {

	public  void main(String[] args) throws Exception {
		userBean bean=new userBean("ram", 25, "telugu");
		userBean bean1=new userBean("hassim", 24, "urdu");
		userBean bean2=new userBean("jack", 22, "english");
		List<userBean> beans=(List) Arrays.asList(bean,bean1,bean2);
		FileOutputStream fout=new FileOutputStream("C:\\Users\\Lenova\\Desktop\\bean.xlsx");
		XSSFWorkbook workbook =new XSSFWorkbook();
		XSSFSheet sheet=workbook.createSheet("beandata");
		Row row=sheet.createRow(0);
		row.createCell(0).setCellValue("name");
		row.createCell(1).setCellValue("age");
		row.createCell(2).setCellValue("language");
		writebean(workbook, sheet, beans);
		
		workbook.write(fout);

	}
	public static void writebean(XSSFWorkbook workbook,XSSFSheet sheet,List<userBean> beans) {
		int rows=1;
	for(userBean bean:beans) {
		Row row=sheet.createRow(rows++);
		row.createCell(0).setCellValue(bean.getName());
		row.createCell(1).setCellValue(bean.getAge());
		row.createCell(2).setCellValue(bean.getLanguage());
		
		
	}
		
	}

}
class userBean {
	String name;
	int age;
	 String language;
	 
	public userBean(String name, int age, String language) {
		this.name = name;
		this.age = age;
		this.language = language;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public int getAge() {
		return age;
	}
	public void setAge(int age) {
		this.age = age;
	}
	public String getLanguage() {
		return language;
	}
	public void setLanguage(String language) {
		this.language = language;
	}
	 
}
