package excel;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataFromExcel {

	public static void main(String[] args) throws IOException {

		FileInputStream file = new FileInputStream(
				"C:\\Users\\jawad\\eclipse-workspace\\SpaceProject\\excel\\data.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheet("Sheet1");

		// to count the rows
		int totalrows = sheet.getLastRowNum();
		int totalcells = sheet.getRow(1).getLastCellNum();

		// to read data
		for (int r = 0; r <= totalrows; r++) {
			XSSFRow rows = sheet.getRow(r);

			for (int c = 0; c < totalcells; c++) {
				XSSFCell cell = rows.getCell(c);
				System.out.print(cell.toString()+"\t");

			}
			System.out.println();
		}
		workbook.close();
		file.close();

	}

}
