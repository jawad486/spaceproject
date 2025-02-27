package excel;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DynamicDataExcel {

	public static void main(String[] args) throws IOException {
		FileOutputStream file1 = new FileOutputStream("C:\\Users\\jawad\\Desktop\\my2.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("data1");

		Scanner sc = new Scanner(System.in);
		System.out.println("enter the rows no");
		int rows = sc.nextInt();

		System.out.println("enter the cells no");
		int cell = sc.nextInt();

		for (int r = 0; r <= rows; r++) {
			XSSFRow currentrow = sheet.createRow(r);

			for (int c = 0; c < cell; c++) {
				XSSFCell currentcell=currentrow.createCell(c);
				currentcell.setCellValue(sc.next());
			}
			}
		workbook.write(file1);
		workbook.close();
		file1.close();
		}
		
	
}