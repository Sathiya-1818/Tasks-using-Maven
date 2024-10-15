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

public class DataintoExcel {

	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet Sheet = wb.createSheet("Sheet1");
		
		Object[][] data = {
				{"   name" , " age " , " Email"},
				
				{"John Doe  " , " 30 " , "  john@test.com"},
				
				{"Jane Doe  " , " 28 " , " john@test.com"},
				
				{"Bob Smith " , "35 " , "  jacky@example.com"},
				
				{"Swapnil   " , " 37 " , "  swapnil@example.com"}
				
		};
		
		int row = 0;
		for(Object[] row1  : data) {
			XSSFRow R = Sheet.createRow(row++);
			
			int column = 0;
			for(Object column1 : row1 ) {
				XSSFCell cell = R.createCell(column++);
				
				if(column1 instanceof String) {
					cell.setCellValue((String)column1);
				} else if(column1 instanceof Integer) {
					cell.setCellValue((Integer)column1);
				}
				
			}
		}
		System.out.println();
		System.out.println();
		System.out.println("Data Entered Succesfully");	
		
		
	try {
		
	FileOutputStream Workbook = new FileOutputStream("C:\\Users\\Public\\eclipsr ptojects\\Fileoperations\\src\\main\\java\\org\\guvi\\Creatingfile.xlsx");
	wb.write(Workbook);
	
	} catch (Exception e) {
		
		e.printStackTrace();
	}
	
	System.out.println();
	System.out.println();
	wb.close();
	
		
	}

}
