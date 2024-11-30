package excel;

import java.io.*;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.chrome.ChromeDriver;

public class ExcelDemo {

	public static void main(String[] args) throws IOException {
		String userName,password;
		ChromeDriver driver=new ChromeDriver();
		FileInputStream foi=new FileInputStream(new File("C://Users//SMILE//Documents/sauce_demo_login.xlsx"));
		//FileOutputStream fos = new FileOutputStream("C://Users//SMILE//Documents/sauce_demo_login.xlsx", true);
		try {
			XSSFWorkbook workbook=new XSSFWorkbook(foi);
			XSSFSheet sheet=workbook.getSheet("Sheet1");
			Iterator<Row> row=sheet.rowIterator();
			Row currRow=row.next();
			while(row.hasNext()) {
				currRow=row.next();
				Iterator<Cell> cell=currRow.cellIterator();
				while(cell.hasNext()) {
					Cell currCell=cell.next();
					userName=currCell.getStringCellValue();
					currCell=cell.next();
					password=currCell.getStringCellValue();
					System.out.print(userName+" "+password+" ");
					Cell newCell=currRow.createCell(2);
					driver.get("https://www.saucedemo.com/");
					driver.findElement(By.id("user-name")).sendKeys(userName);
					driver.findElement(By.id("password")).sendKeys(password);
					driver.findElement(By.id("login-button")).click();
					String URL = driver.getCurrentUrl();
					if(URL.equalsIgnoreCase("https://www.saucedemo.com/inventory.html" )) {
						System.out.println("login successful");
						newCell.setCellValue("login successful");
					}
					else {
						System.out.println("login not successful");
						newCell.setCellValue("login not successful");
					}
					//workbook.write(fos);
				}
			}
			workbook.close();
			driver.close();
			
		} catch (IOException e) {
			
			e.printStackTrace();
		}

	}

}
