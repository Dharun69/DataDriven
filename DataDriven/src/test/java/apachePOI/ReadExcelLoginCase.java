package apachePOI;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.time.Duration;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class ReadExcelLoginCase {

	public static void main(String[] args) throws IOException {

		WebDriver driver = new ChromeDriver();
		driver.navigate().to("http://demowebshop.tricentis.com/login");
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));

		
		String file = "E:\\Excel\\Demologin.xlsx";
		FileInputStream excel = new FileInputStream(file);

		XSSFWorkbook workbook = new XSSFWorkbook(excel);
		XSSFSheet sheet =	workbook.getSheet("Sheet1");

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
			//System.out.println("Username is :"+ username + "  "+ "Password is :"+ password);

			WebElement email = driver.findElement(By.xpath("//input[@id='Email']"));
			email.sendKeys(username);
			WebElement pw = driver.findElement(By.xpath("//input[@id='Password']"));
			pw.sendKeys(password);

			WebElement login = driver.findElement(By.xpath("//input[@type='submit'] [@value='Log in']"));
			login.click();

			String result = null;
			try
			{
				boolean isDisplayed =	driver.findElement(By.xpath("//a[@class='ico-logout']")).isDisplayed();
				if(isDisplayed==true)
				{
					result = "Pass";
					//Writing to an excel
					cell=row.createCell(2);
					cell.setCellType(CellType.STRING);
					cell.setCellValue(result);
				}
				System.out.println("Username is :"+ username + "  "+ "Password is :"+ password + "Is login success "+ result);
				driver.findElement(By.xpath("//a[@class='ico-logout']")).click();

			}
			catch(Exception e)
			{
				boolean isError = driver.findElement(By.xpath("//*[text()='The credentials provided are incorrect']")).isDisplayed();
				if(isError==true) {
					result="Fail";
					cell=row.createCell(2);
					cell.setCellType(CellType.STRING);
					cell.setCellValue(result);
				}
				System.out.println("Username is :"+ username + "  "+ "Password is :"+ password + "Is login success "+ result);

		
			}
			Thread.sleep(1000);

			driver.findElement(By.xpath("//a[@class='ico-login']")).click();


		}
                FileOutputStream fileOutputStream = new FileOutputStream(file);
		workbook.write(fileOutputStream);
		fileOutputStream.close();
	}
}
