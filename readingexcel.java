package excelOperations;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.*;

public class readingexcel {
	
	public static void main(String args[]) throws IOException {
		
		FileInputStream inputStream = new FileInputStream(".\\data\\file1.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		
		int rows = sheet.getLastRowNum();
		int cols = sheet.getRow(1).getLastCellNum();
		
		for(int r=0;r<rows;r++) {
			XSSFRow row = sheet.getRow(r);
			for(int c=0;c<cols;c++)
			{
				XSSFCell cell = row.getCell(c);
				switch(cell.getCellType()) {
				case STRING: System.out.print(cell.getStringCellValue()+"\t");
							break;
				
				case NUMERIC: System.out.print(cell.getNumericCellValue()+"\t");
							break;
				}
			}
			System.out.println();
		}
	}
	
}
