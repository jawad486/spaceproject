package excel;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PractReadingData {

	public static void main(String[] args) throws IOException {

		FileInputStream file = new FileInputStream("C:\\Users\\jawad\\Desktop\\sec.xlsx");

		// creating workbook
		XSSFWorkbook workbook = new XSSFWorkbook(file);

		// creating sheet

		XSSFSheet sheet = workbook.getSheet("Data");

		XSSFRow row = sheet.getRow(0);
		XSSFCell cell = row.getCell(2);
		System.out.println(cell);
		System.out.println("====================================================================");

		// how to read all the data from excel

		int totalrows = sheet.getLastRowNum();
		int totalcells = sheet.getRow(0).getLastCellNum();
		
		for(int r=0;r<=totalrows;r++) {
			   XSSFRow currentrow=  sheet.getRow(r);
			for(int c=0;c<totalcells;c++) {
				
				XSSFCell cell1 =currentrow.getCell(c);
				System.out.print(cell1.toString()+"\t");
				
				
			}
			System.out.println();
		}

	}

}
