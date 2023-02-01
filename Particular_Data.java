package com.data;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Particular_Data {

	public static void main(String[] args) throws IOException {

		File f = new File("C:\\Users\\Lenovo\\eclipse-workspace\\Project_Morning\\ReadWrite\\Read.xlsx");

		FileInputStream fis = new FileInputStream(f);

		Workbook wb = new XSSFWorkbook(fis);

		Sheet s = wb.getSheet("Data");

		Row r = s.getRow(1);

		Cell c = r.getCell(2);

		CellType type = c.getCellType(); // String /Numeric

		if (type.equals(CellType.STRING)) {

			String value = c.getStringCellValue();
			System.out.println(value);

		} else if (type.equals(CellType.NUMERIC)) {

			double n = c.getNumericCellValue();

			int i = (int) n;

			String value = String.valueOf(i);

			System.out.println(value);

		}

	
		
	
	}

}
