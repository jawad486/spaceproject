package excel;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingUserInput {

	
	public static void main (String [] args) throws IOException {
		FileOutputStream file = new FileOutputStream("C:\\Users\\jawad\\Desktop\\sec1.xlsx");
		XSSFWorkbook workbook=new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("pagal");
		
		Scanner scanner=new Scanner(System.in);
		System.out.println("create the rows");
		int rows=scanner.nextInt();
		
		System.out.println("create the cell");
		int cells=scanner.nextInt();
		
		for(int r=0;r<=rows;r++) {
			XSSFRow rerow = sheet.createRow(r);
			
			for(int c=0;c<cells;c++) {
				XSSFCell ce = rerow.createCell(c);
				ce.setCellValue(scanner.nextInt());
				
			}
		}
		workbook.write(file);
		workbook.close();
	}
}
