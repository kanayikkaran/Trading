package com.myauto;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.UnhandledAlertException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

/**
 * Scrip Finder
 */
public class CopyOfScripFinder
{
	public static String filepath = "./src/main/resources/";
	public static String filename = "BHAV_05JUL2017";
	public static String prevfilename = "BHAV_23JUN2017";
	
	public static double cut_off_percent = 0.5;
	
	public static WebDriver driver;
	public static String file = filepath+filename+".xlsx";
	public static String outputfilename = filepath+filename+".txt";
	
	@BeforeMethod
	public void loadPage() throws Exception
	{
		System.setOut(new PrintStream(new FileOutputStream(outputfilename)));
		System.out.println("Finding Scrip's Delivery%");
		File file = new File(".//src//main/drivers//chromedriver.exe");
		System.setProperty("webdriver.chrome.driver", file.getAbsolutePath());
		driver = new ChromeDriver();
		driver.get("https://www.nseindia.com/index_nse.htm");
		driver.manage().window().maximize();
	}
	
	@AfterMethod
	public void closePage()
	{
		driver.quit();
	}
	
	@Test(priority=1, enabled=false)
	public void updateDiff()
	{
		updateDifference(file);
	}
	
	@Test(priority=2, enabled=true)
	public void checkDelivery()
	{
		try {
		List<String> scrips = getScripName(file);
		for(int i=0; i<scrips.size(); i++)
		{
			try
			{
			String item = scrips.get(i);
			driver.get("https://www.nseindia.com/live_market/dynaContent/live_watch/get_quote/GetQuote.jsp?symbol="+item+"&illiquid=0&smeFlag=0&itpFlag=0");
			WebElement ele = driver.findElement(By.xpath("//h5[contains(.,'Security-wise Delivery Position')]"));
			Thread.sleep(1000*1);
			scrollUp(250);
			ele.click();
			Thread.sleep(1000*1);
			String delvalue = driver.findElement(By.xpath(".//*[@id='deliveryToTradedQuantity']")).getText();
			if(delvalue.length()>1)
			{
				setDeliveryValue(i+1, Double.valueOf(delvalue));
			}
			
			}
			catch(UnhandledAlertException e)
			{
				driver.switchTo().alert().accept();
			}
		}
		} catch (Exception e) {
			
			e.printStackTrace();
		}
	}
	
	@Test(priority=3, enabled=false)
	public void compareScrips()
	{
		List<String> common = findCommon(filepath+filename+".xlsx", filepath+prevfilename+".xlsx");
		System.out.println("Common Scrips Found : "+common.size());
		System.out.println(common);
	}
	
	public List<String> findCommon(String file1, String file2)
	{
		List<String> scrips1 = getScripName(file1);
		List<String> scrips2 = getScripName(file2);
		System.out.println(filename+"_Size="+scrips1.size());
		System.out.println(prevfilename+"_Size="+scrips2.size());
		List<String> commonScrips = new ArrayList<String>();
		for(String item:scrips1)
			if(scrips2.contains(item))
				commonScrips.add(item);
		return commonScrips;
	}
	
	public List<String> getScripName(String file)
	{
		List<String> scrip = new ArrayList<String>();
		
		try {
			FileInputStream fis = null;
			XSSFWorkbook workbook = null;
			XSSFSheet sheet = null;
			XSSFRow row = null;
			XSSFCell changepCell = null;
			XSSFCell cell = null;
			fis = new FileInputStream(file);
			workbook = new XSSFWorkbook(fis);
			sheet = workbook.getSheetAt(0);
			fis.close();
			
			sheet = workbook.getSheetAt(0);
			int symbolVal=0;
			row=sheet.getRow(0);
			for(int i=0;i<row.getLastCellNum();i++)
			{
				if(row.getCell(i).getStringCellValue().trim().equalsIgnoreCase("SYMBOL"))
					symbolVal=i;
			}
			
			int changepVal=0;
			for(int i=0;i<row.getLastCellNum();i++)
			{
				if(row.getCell(i).getStringCellValue().trim().equalsIgnoreCase("CHANGEPERCENT"))
					changepVal=i;
			}
			
			for(int i=1; i<sheet.getPhysicalNumberOfRows(); i++)
			{
			row=sheet.getRow(i);
			cell = row.getCell(symbolVal);
			changepCell = row.getCell(changepVal);
			if(changepCell.getNumericCellValue()>cut_off_percent)
				scrip.add(cell.getStringCellValue());
			}
			workbook.close();
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return scrip;
	}
	
	public void updateDifference(String file)
	{
		
		try {
			FileInputStream fis = null;
			FileOutputStream fos = null;
			XSSFWorkbook workbook = null;
			XSSFSheet sheet = null;
			XSSFRow row = null;
			XSSFCell cell = null;
			fis = new FileInputStream(file);
			workbook = new XSSFWorkbook(fis);
			sheet = workbook.getSheetAt(0);
			fis.close();
			row = sheet.getRow(0);
			int closeVal=0;
			int prevCloseVal=0;
			for(int i=0;i<row.getLastCellNum();i++)
			{
				if(row.getCell(i).getStringCellValue().trim().equalsIgnoreCase("CLOSE"))
					closeVal=i;
				if(row.getCell(i).getStringCellValue().trim().equalsIgnoreCase("PREVCLOSE"))
					prevCloseVal=i;
			}
			
			row.createCell(row.getLastCellNum());
			row.getCell(row.getLastCellNum()-1).setCellValue("CHANGE");
			int diffVal = row.getLastCellNum()-1;
			
			row.createCell(row.getLastCellNum());
			row.getCell(row.getLastCellNum()-1).setCellValue("CHANGEPERCENT");
			int diffpVal = row.getLastCellNum()-1;
			
			row.createCell(row.getLastCellNum());
			row.getCell(row.getLastCellNum()-1).setCellValue("DELIVERY");
			
			fos = new FileOutputStream(file);
			workbook.write(fos);
			fos.close();
			
			for(int i=1; i<sheet.getPhysicalNumberOfRows(); i++)
			{
			row=sheet.getRow(i);
			XSSFCell closeCell = row.getCell(closeVal);
			XSSFCell prevcloseCell = row.getCell(prevCloseVal);
			
			cell = row.getCell(diffVal);
			if(cell==null)
				cell = row.createCell(diffVal);
			cell.setCellValue(closeCell.getNumericCellValue()-prevcloseCell.getNumericCellValue());
			
			cell = row.getCell(diffpVal);
			if(cell==null)
				cell = row.createCell(diffpVal);
			cell.setCellValue((closeCell.getNumericCellValue()-prevcloseCell.getNumericCellValue())/prevcloseCell.getNumericCellValue()*100);
			}
			
			FormulaEvaluator fe = workbook.getCreationHelper().createFormulaEvaluator();
			fe.evaluateAll();
			fos = new FileOutputStream(file);
			workbook.write(fos);
			workbook.close();
			fos.close();
		
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
	}
	
	public void setDeliveryValue(int rowNum, String value)
	{
		try {
			FileInputStream fis = null;
			FileOutputStream fos = null;
			XSSFWorkbook workbook = null;
			XSSFSheet sheet = null;
			XSSFRow row = null;
			XSSFCell cell = null;
			fis = new FileInputStream(file);
			workbook = new XSSFWorkbook(fis);
			sheet = workbook.getSheetAt(0);
			row = sheet.getRow(0);
			int colNum=0;
			boolean deliveryfound=false;
			for(int i=0;i<row.getLastCellNum();i++)
			{
				if(row.getCell(i).getStringCellValue().trim().equalsIgnoreCase("DELIVERY"))
					{
					colNum=i;
					deliveryfound=true;
					}
			}
			if(!deliveryfound)
			{
				sheet.getRow(0).createCell(sheet.getRow(0).getLastCellNum()).setCellValue("DELIVERY");
				colNum = sheet.getRow(0).getLastCellNum()-1;
			}
			row = sheet.getRow(rowNum);
			cell = row.getCell(colNum);
			if(cell==null)
				cell = row.createCell(colNum);
			cell.setCellValue(value);
			
			fis.close();
			fos = new FileOutputStream(file);
			workbook.write(fos);
			workbook.close();
			fos.close();
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public void setDeliveryValue(int rowNum, double value)
	{
		try {
			FileInputStream fis = null;
			FileOutputStream fos = null;
			XSSFWorkbook workbook = null;
			XSSFSheet sheet = null;
			XSSFRow row = null;
			XSSFCell cell = null;
			fis = new FileInputStream(file);
			workbook = new XSSFWorkbook(fis);
			sheet = workbook.getSheetAt(0);
			row = sheet.getRow(0);
			int colNum=0;
			boolean deliveryfound=false;
			for(int i=0;i<row.getLastCellNum();i++)
			{
				if(row.getCell(i).getStringCellValue().trim().equalsIgnoreCase("DELIVERY"))
					{
					colNum=i;
					deliveryfound=true;
					}
			}
			if(!deliveryfound)
			{
				sheet.getRow(0).createCell(sheet.getRow(0).getLastCellNum()).setCellValue("DELIVERY");
				colNum = sheet.getRow(0).getLastCellNum()-1;
			}
			row = sheet.getRow(rowNum);
			cell = row.getCell(colNum);
			if(cell==null)
				cell = row.createCell(colNum);
			cell.setCellValue(value);
			
			fis.close();
			fos = new FileOutputStream(file);
			workbook.write(fos);
			workbook.close();
			fos.close();
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public static void waitForPageToLoad() {
		String pageLoadStatus;
		do {
			JavascriptExecutor js = (JavascriptExecutor) driver;
			pageLoadStatus = (String)js.executeScript("return document.readyState");
		} while ( !pageLoadStatus.equals("complete") );
	}

	public static void scrollUp(int value){
		JavascriptExecutor js = ((JavascriptExecutor) driver);
		js.executeScript("scroll(0, "+value+");");
	}
	
}
