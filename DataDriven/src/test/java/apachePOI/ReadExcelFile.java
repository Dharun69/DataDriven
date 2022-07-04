package apachePOI;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelFile {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		//	File file = new File(System.getProperty("user.dir") + "\\TestData\\"+ "demodatadriven"+ ".xlsx");

		String file = "E:\\Excel\\demodatadriven.xlsx";
		FileInputStream excel = new FileInputStream(file);

		XSSFWorkbook workbook = new XSSFWorkbook(excel);
		XSSFSheet sheet =	workbook.getSheet("Sheet1");
        XSSFRow row = sheet.getRow(1);
        XSSFCell cell = row.getCell(0); //get 1st username
        
        String username = cell.getStringCellValue();
        System.out.println("Username is :"+username);

        cell = row.getCell(1);
        String password = cell.getStringCellValue();
        System.out.println("Password is :"+password);

	}

}
