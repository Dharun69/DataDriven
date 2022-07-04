package apachePOI;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelUsingColumnNames {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		//	File file = new File(System.getProperty("user.dir") + "\\TestData\\"+ "demodatadriven"+ ".xlsx");

		String file = "E:\\Excel\\demodatadriven.xlsx";
		FileInputStream excel = new FileInputStream(file);

		XSSFWorkbook workbook = new XSSFWorkbook(excel);
		XSSFSheet sheet =	workbook.getSheet("Sheet1");
        // Getting number of columns , and getting the column(header) names
		// Printing allthe values from the column
		// Storing values in to variables to pass to the web application for testing
		
		
		//uncomment below code is // Getting number of columns , and getting the column(header) names
		/*
		 * XSSFRow row = sheet.getRow(0);
		 * 
		 * //getting the column count 
		 * int columnCount = row.getLastCellNum();
		 * System.out.println("Column Count is "+ columnCount);
		 * 
		 * 
		 * XSSFCell cell = null;
		 * 
		 * //getting the column(headers) names 
		 * for(int i=0; i<columnCount;i++) {
		 * cell=row.getCell(i); 
		 * String column = cell.getStringCellValue();
		 * System.out.println("Column name is "+ column); }
		 */
		
		
		//Printing allthe values from the column
		
		/*XSSFRow row= null;
		XSSFCell cell= null;
		
		for(int i=0; i<=sheet.getLastRowNum(); i++) // navigate through the row
		{
			row= sheet.getRow(i);
			
			for(int j=0; j<row.getLastCellNum();j++)  // navigate through the columns
			{
				cell=row.getCell(j);
				String myCellValue = cell.getStringCellValue();
				System.out.println("My cell value is "+ myCellValue);
			}
			
		}*/
		
		
		//Storing values in to variables (in this case, username and password)
		
		XSSFRow row = null;
		XSSFCell cell = null;
		String username = null;
		String password = null;
				
		for(int i=1; i<=sheet.getLastRowNum(); i++)
		{
			row=sheet.getRow(i);
			
			for(int j=0; j<row.getLastCellNum(); j++) {
				cell= row.getCell(j);
			
				if(j==0) // we can use column name as well.
				{
					username= cell.getStringCellValue();
				}
				
				if(j==1)  //we can use colum name as well
				{
					password = cell.getStringCellValue();
				}
				
			}
			System.out.println("Username is :"+ username + "  "+ "Password is :"+ password);

			//we can pass this values in to web application for testing test user accounts
		}

	}

}
