package com.validlog.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.ITestResult;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

public class SendDetailsTest {
	WebDriver driver;
	
	File file = null;
	FileInputStream fis = null;
	Sheet sheet = null;
	Workbook workbook = null;
	FileOutputStream fos = null;
	String fileExtension = null;

	@BeforeMethod
	public void initialization() {
		System.setProperty("webdriver.chrome.driver", "D:\\Testing\\Selenium\\ValidLogs\\driver\\chromedriver.exe");

		driver = new ChromeDriver();

		driver.manage().window().maximize();
		driver.manage().deleteAllCookies();
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		driver.manage().timeouts().pageLoadTimeout(60, TimeUnit.SECONDS);
		driver.get("https://www.validlog.com/contact-us/");
	}

	@Test
	public void check() throws Exception {
		String fileName = "D:\\Testing\\Selenium\\ValidLogs\\src\\com\\validlog\\test\\details.xlsx";
		file = new File(fileName);
		fis = new FileInputStream(file);
		fileExtension = fileName.substring(fileName.indexOf("."));

		if (fileExtension.equals(".xlsx")) {
			workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook(fis);
		} else if (fileName.equals(".xls")) {
			workbook = new HSSFWorkbook(fis);
		}
		sheet = workbook.getSheetAt(0);
		int rowNum = sheet.getLastRowNum();
		HashMap<String, String> hashtable = new HashMap<String, String>();
		for (int i = 0; i <= rowNum; i++) {
			Row row = sheet.getRow(i);
			for (int j = 0; j < row.getLastCellNum(); j++) {
				String header = sheet.getRow(i).getCell(j).getStringCellValue();
				String value = sheet.getRow(i).getCell(1).getStringCellValue();
				hashtable.put(header, value);
				break;
			}
		}
		String name = hashtable.get("Name");
		String email = hashtable.get("Email");
		String phone = hashtable.get("Phone");
		String message = hashtable.get("Message");
		
		WebElement scroll = driver
				.findElement(By.xpath("//div[@class='elementor-widget-container']/h5[contains(text(),'KEEP')]"));
		JavascriptExecutor js = (JavascriptExecutor) driver;
		js.executeScript("arguments[0].scrollIntoView();", scroll);

		WebElement nameField = driver.findElement(By.xpath("//input[@id='form-field-name']"));
		WebElement phoneField = driver.findElement(By.xpath("//input[@id='form-field-phone']"));
		WebElement emailField = driver.findElement(By.xpath("//input[@id='form-field-email']"));
		WebElement messageField = driver.findElement(By.xpath("//textarea[@id='form-field-message']"));
		WebElement clickButton = driver.findElement(By.xpath("//button//span[2]"));
		nameField.sendKeys(name);
		phoneField.sendKeys(phone);
		emailField.sendKeys(email);
		messageField.sendKeys(message);
		clickButton.click();

	}

	@AfterMethod
	public void tearDown(ITestResult result) throws Exception {
		String fileName = "D:\\Testing\\Selenium\\ValidLogs\\src\\com\\validlog\\test\\details.xlsx";
		if (result.getStatus() == ITestResult.SUCCESS) {
			file = new File(fileName);
			fis = new FileInputStream(file);
			fileExtension = fileName.substring(fileName.indexOf("."));
			if (fileExtension.equals(".xlsx")) {
				workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook(fis);
			} else if (fileName.equals(".xls")) {
				workbook = new HSSFWorkbook(fis);
			}
			sheet = workbook.getSheetAt(0);
			int rowNum = sheet.getLastRowNum();
			Row row = sheet.createRow(rowNum + 1);
			Cell cell = row.createCell(0);
			cell.setCellValue("Status");
			Cell cell1 = row.createCell(1);
			cell1.setCellValue("PASSED");
			fos = new FileOutputStream(file);
			workbook.write(fos);
			fos.close();
		} else {
			
			file = new File(fileName);
			fis = new FileInputStream(file);
			fileExtension = fileName.substring(fileName.indexOf(","));
			if (fileExtension.equals(".xlsx")) {
				workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook(fis);
			} else if (fileName.equals(".xls")) {
				workbook = new HSSFWorkbook(fis);
			}
			sheet = workbook.getSheetAt(0);
			int rowNum = sheet.getLastRowNum();
			Row row = sheet.createRow(rowNum + 1);
			Cell cell = row.createCell(0);
			cell.setCellValue("Status");
			Cell cell1 = row.createCell(1);
			cell1.setCellValue("FAILED");
			fos = new FileOutputStream(file);
			workbook.write(fos);
			fos.close();
			
		}

	}

}
