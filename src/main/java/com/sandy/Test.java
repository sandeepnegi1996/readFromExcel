package com.sandy;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class Test {

	public static void main(String[] args) throws FileNotFoundException, IOException {

		writeToExcel("TestModes.xlsx");
		System.out.println("writing is done");
		readFromExcel("TestModes.xlsx");
	}

	public static void readFromExcel(String file) throws FileNotFoundException, IOException {

		HSSFWorkbook myexcelbook = new HSSFWorkbook(new FileInputStream(file));

		HSSFSheet myexcelSheet = myexcelbook.getSheet("TestCases");

		
		  HSSFRow row = myexcelSheet.getRow(0);
		  
		  String testname = row.getCell(0).getStringCellValue();
		  System.out.println(testname);
		  
		  HSSFRow row1 = myexcelSheet.getRow(1);
		  
		  String testname1 = row1.getCell(0).getStringCellValue();
		  System.out.println(testname1);
		  
		 
		
		
		

		myexcelbook.close();

	}

	public static void writeToExcel(String file) {

		HSSFWorkbook book = new HSSFWorkbook();
		HSSFSheet sheet = book.createSheet("TestCases");

		// first row
		Row row = sheet.createRow(0);

		Cell name = row.createCell(0);
		name.setCellValue("APGARequestCreation");

		Row row1 = sheet.createRow(1);
		Cell test2 = row1.createCell(0);
		test2.setCellValue("RBGA Reqeust Creation");

		try {
			book.write(new FileOutputStream(file));

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		try {
			book.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

}
