package com.ExcelTest;

import org.testng.annotations.Test;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import org.testng.annotations.BeforeMethod;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.DataProvider;
import org.testng.annotations.BeforeClass;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.AfterSuite;

public class ValidationwithExcel {
	
	WebDriver driver =null;
	
	@BeforeTest
	public void BeforeTest() 
	{
		System.setProperty("webdriver.chrome.driver", "src/test/resources/Driver/chromedriver");
		driver =new ChromeDriver();
		
		driver.get("https://www.google.com");
		
		driver.manage().timeouts().implicitlyWait(8, TimeUnit.SECONDS);
		driver.manage().window().maximize();
		
	}
	
	@Test(dataProvider= "ReadExcel")
	public void GoogleTest(String cellvalue1, String cellvalue2) throws InterruptedException
	{
		
		
		//if (cellvalue1.equalsIgnoreCase("Y"))
		//	{
		    driver.findElement(By.name("q")).sendKeys(cellvalue2);
			Thread.sleep(6000);
			driver.findElement(By.name("q")).clear();
		//}
		//	else
		//		
			//	System.out.println("Skipped");
			
	}
	
			
	@DataProvider(name = "ReadExcel")
	public String[][] readTestData() throws BiffException, IOException, NullPointerException
	{
	 String[][] content = null; 
	 			 
		 FileInputStream fis = null;
		try {
			fis = new FileInputStream("src/test/resources/TestData/TD.xls");
		} catch (FileNotFoundException e) {
			
			e.printStackTrace();
		}
		 Workbook wb = Workbook.getWorkbook(fis);
		 Sheet sh = wb.getSheet("Sheet1");
		 
		 int totalNoOfRows = sh.getRows();
		 int totalNoOfColumns = sh.getColumns();
		 
		 System.out.println("Value of Columns  is : "+ totalNoOfColumns);
		 System.out.println("Value of Rows  is : "+ totalNoOfRows);
		 
		 content = new String[totalNoOfRows][totalNoOfColumns];
		 	 
		 int i=0;
		 int j=0;
		 for( i =0;i < totalNoOfRows;i++)
		 {     
			 for( j =0;j < totalNoOfColumns;j++)
			 {
				 Cell cell =sh.getCell(j, i);
				 content [i][j]= cell.getContents();
				System.out.println("Data is  i-->" + i + "  j-->"+ j + "  " + content[i][j]);
		
		
		if (i<=totalNoOfRows && content[i][j].startsWith("N"))
			{
				System.out.println("The Content is  i-->" + i + "  j-->"+ j + "  " + content[i][j]);
				i++;
				j=0;
			}	
		 
		
		
		 }
			 
		 
		 }
		
		return content;
		 
		 
		
	}
		 
	
	
 
  @AfterTest
	public void browserquit()
	{
		driver.quit();
		
	}
	
	

}
  



	