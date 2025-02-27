package excel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteDataInExcelSheet {

	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Data");
		XSSFRow row1 = sheet.createRow(0);
		row1.createCell(0).setCellValue("Name");
		row1.createCell(1).setCellValue("location");
		row1.createCell(2).setCellValue("occupation");

		XSSFRow row2 = sheet.createRow(1);
		row2.createCell(0).setCellValue("jawad");
		row2.createCell(1).setCellValue("New jersey");
		row2.createCell(2).setCellValue("Tester");

		XSSFRow row3 = sheet.createRow(2);
		row3.createCell(0).setCellValue("shahid");
		row3.createCell(1).setCellValue(" Doaba");
		row3.createCell(2).setCellValue("Teacher");

		XSSFRow row4 = sheet.createRow(3);
		row4.createCell(0).setCellValue("Farhan");
		row4.createCell(1).setCellValue("Doaba");
		row4.createCell(2).setCellValue("Admin");

		XSSFRow row5 = sheet.createRow(4);
		row5.createCell(0).setCellValue("faraz");
		row5.createCell(1).setCellValue("Qatar");
		row5.createCell(2).setCellValue("Eingeeer");
		FileOutputStream file1 = new FileOutputStream(
				"C:\\Users\\jawad\\Desktop\\sec.xlsx");
		workbook.write(file1);
		workbook.close();
		file1.close();
		
		System.out.println("file is created");

	}

}
