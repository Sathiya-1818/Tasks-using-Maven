package org.guvi;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook wb = new XSSFWorkbook("C:\\Users\\Public\\eclipsr ptojects\\Fileoperations\\src\\main\\java\\org\\guvi\\Creatingfile.xlsx");
		
		XSSFSheet sheet = wb.getSheet("Sheet1");
		
		int row_coun = sheet.getLastRowNum();
		
		int column_coun = sheet.getRow(0).getLastCellNum();
		
		for (int i = 1; i <= row_coun; i++) {
			XSSFRow row = sheet.getRow(i);
			
			for (int j = 0; j < column_coun; j++) {
				XSSFCell cell = row.getCell(j);
				
				System.out.print(cell.getStringCellValue()+"  ");
				
			}
			
			System.out.println();	
			
		}

	}

}
