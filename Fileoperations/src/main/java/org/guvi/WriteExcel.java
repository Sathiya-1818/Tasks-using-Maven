package org.guvi;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {

	public static void main(String[] args) throws IOException {
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet Sheet = wb.createSheet("Sheet1");
		
		Object[][] data = {
				{"name" , "age","Email"},
				
		};
		
		int row = 0;
		for(Object[] row1  : data) {
			XSSFRow R = Sheet.createRow(row++);
			
			int column = 0;
			for(Object column1 : row1 ) {
				XSSFCell cell = R.createCell(column++);
				
			}
		}
		System.out.println();
		System.out.println();
		System.out.println("Sheet Created Succesfully");	
		
	try {	
	FileOutputStream Workbook = new FileOutputStream("C:\\Users\\Public\\eclipsr ptojects\\Fileoperations\\src\\main\\java\\org\\guvi\\Creatingfile.xlsx");
	wb.write(Workbook);
	} catch (Exception e) {
		
		e.printStackTrace();
	}
	wb.close();
	
		
	}

}
