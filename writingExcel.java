package excelOperations;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.*;

public class writingExcel {

public static void main(String args[]) throws IOException {
		
		@SuppressWarnings("resource")
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("creating sheet");
		
		Object[][] empdata = {
								{"name","age"},
								{"Rossi","22"},
								{"Jessy","15"},
								{"Messi","17"}
							 };
		
		int rows= empdata.length;
		int cols= empdata[0].length;
		
		for(int r=0;r<rows;r++)
		{
			XSSFRow row = sheet.createRow(r);
			
			for(int c=0;c<cols;c++)
			{
				XSSFCell cell = row.createCell(c);
				Object val = empdata[r][c];	
				
				if(val instanceof String) 
					cell.setCellValue((String)val);;
				if(val instanceof Integer) 
					cell.setCellValue((Integer)val);;
			}
		}
		
		FileOutputStream outputStream=new FileOutputStream(".\\data\\emp.xlsx");
		workbook.write(outputStream);
		outputStream.close();
		System.out.println("Created a new excell sheet ...");
		
	}
	
}
