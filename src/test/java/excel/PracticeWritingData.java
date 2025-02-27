package excel;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PracticeWritingData {

	public static void main(String[] args) throws IOException {

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("data");

		XSSFRow row1 = sheet.createRow(0);
		row1.createCell(0).setCellValue("ID");
		row1.createCell(1).setCellValue("name");
		row1.createCell(2).setCellValue("salary");

		XSSFRow row2 = sheet.createRow(1);
		row2.createCell(0).setCellValue(005);
		row2.createCell(1).setCellValue("jay");
		row2.createCell(2).setCellValue(3500.68);
		FileOutputStream file1 = new FileOutputStream("C:\\Users\\jawad\\Desktop\\my.xlsx");
		workbook.write(file1);
		workbook.close();
		file1.close();
		System.out.println("file is create");

	}

}
