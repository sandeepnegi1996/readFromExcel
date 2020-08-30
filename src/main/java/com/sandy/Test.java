package com.sandy;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test {

	public static void main(String[] args) throws FileNotFoundException, IOException {

		/*
		 * writeToExcel("TestModes.xlsx"); System.out.println("writing is done");
		 * readFromExcel("TestModes.xlsx");
		 */

		// readMultipleRows("TestModes.xlsx");
		writeMultipleRows("TestModes.xlsx");
	}

	public static void writeMultipleRows(String file) throws IOException {

		XSSFWorkbook workbook = new XSSFWorkbook();

		XSSFSheet sheet = workbook.createSheet("EmployeeData");

		Map<String, Object[]> data = new TreeMap<String, Object[]>();

		data.put("1", new Object[] { "ID", "NAME", "LASTNAME" });
		data.put("2", new Object[] { 1, "Amit", "yadav" });
		data.put("3", new Object[] { 2, "Dhruv", "rana" });
		data.put("4", new Object[] { 3, "Vikas", "tiwari" });
		data.put("5", new Object[] { 4, "Alok", "kumar" });

		Set<String> keyset = data.keySet();
		int rowNum = 0;
		for (String key : keyset) {

			Row row = sheet.createRow(rowNum++);

			Object[] objArr = data.get(key);

			int cellnum = 0;

			for (Object obj : objArr) {

				Cell cell = row.createCell(cellnum++);

				if (obj instanceof String) {
					cell.setCellValue((String) obj);
				}

				else if (obj instanceof Integer) {
					cell.setCellValue((Integer) obj);
				}

			}

		}

		try {

			FileOutputStream out = new FileOutputStream(file);
			workbook.write(out);
			out.close();
			System.out.println("Sucessfully written in the file");

		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	public static void readMultipleRows(String file) throws FileNotFoundException, IOException {

		XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(file));

		XSSFSheet sheet = workbook.getSheet("TestCases");

		// iterate through each rows of the sheet
		Iterator<Row> rowIterator = sheet.iterator();

		while (rowIterator.hasNext()) {

			// this is our first row
			Row row = rowIterator.next();

			Iterator<Cell> cellIterator = row.cellIterator();

			while (cellIterator.hasNext()) {

				Cell cell = cellIterator.next();

				switch (cell.getCellType()) {

				case Cell.CELL_TYPE_STRING:
					System.out.println(cell.getStringCellValue() + "  t");
					break;

				case Cell.CELL_TYPE_NUMERIC:
					System.out.println(cell.getNumericCellValue() + "  t");
					break;

				}

			}

			System.out.println(" ");

		}

	}

	public static void readFromExcel(String file) throws FileNotFoundException, IOException {

		XSSFWorkbook myexcelbook = new XSSFWorkbook(new FileInputStream(file));

		XSSFSheet myexcelSheet = myexcelbook.getSheet("TestCases");

		XSSFRow row = myexcelSheet.getRow(0);

		String testname = row.getCell(0).getStringCellValue();
		System.out.println(testname);

		XSSFRow row1 = myexcelSheet.getRow(1);

		String testname1 = row1.getCell(0).getStringCellValue();
		System.out.println(testname1);

		myexcelbook.close();

	}

	public static void writeToExcel(String file) {

		XSSFWorkbook book = new XSSFWorkbook();
		XSSFSheet sheet = book.createSheet("TestCases");

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

			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		try {
			book.close();
		} catch (IOException e) {

			e.printStackTrace();
		}

	}

}
