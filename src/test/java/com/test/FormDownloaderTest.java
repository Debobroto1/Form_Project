package com.test;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

public class FormDownloaderTest {
	private WebDriver driver;
	private WebDriverWait wait;
	private Workbook workbook;

	@BeforeClass
	public void setup() {
		// setting ChromeDriver executable
		System.setProperty("webdriver.chrome.driver", System.getProperty("user.dir")+"\\src\\test\\resources\\executables\\chromedriver.exe");
		// Set Chromeoptions to disable notifications and maximize window
		ChromeOptions options = new ChromeOptions();
		options.addArguments("--disable-notifications");
		options.addArguments("--start-maximized");
		// Create a new instance of the Chrome driver
		driver = new ChromeDriver(options);
		wait = new WebDriverWait(driver, 10);
		try {
			// Load the Excel workbook
			workbook = new XSSFWorkbook(System.getProperty("user.dir")+"//src//test//resources//excel//Document_download_selenium.xlsx");
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	@Test
	public void testFormDownload() throws InterruptedException {
		// Open the link
		driver.get("http://houseofforms.org/ActWiseForms.aspx?Search=&ActID=9650A6D7E5A396D0");

		// Get the first sheet
		Sheet sheet = workbook.getSheetAt(0);

		// Get the total number of rows with data in the sheet			
		int totalRows = sheet.getLastRowNum() + 1;

		// Iterate through each row
		for (int row = 1; row < totalRows; row++) {
			// Get the form name from the current row
			Row currentRow = sheet.getRow(row);
			Cell formNameCell = currentRow.getCell(1);
			String formName = formNameCell.getStringCellValue().replaceAll("[^\\x20-\\x7E]", "").trim();
			// Search the form name on the website search section
			WebElement searchInput = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ctl00_txtSearch")));
			searchInput.clear();
			searchInput.sendKeys(formName);

			WebElement searchButton = driver.findElement(By.id("ctl00_Search"));
			searchButton.click();
			// Capture the start time
			Cell startTimeCell = currentRow.createCell(4);
			startTimeCell.setCellValue(getCurrentDateTime());
			
			// Find all document links and click on each one
			List<WebElement> viewElements = driver.findElements(By.xpath("//a[text()=' View ']"));
			for (WebElement documentLink : viewElements) {
				documentLink.click();
				// Download the form				
				Thread.sleep(3000);
				downloadAndSave();
				// Go back to the previous page for the next document
				driver.navigate().back();
			}
			// Update the status on the Excel sheet
			Cell statusCell = currentRow.createCell(3);
			statusCell.setCellValue("Downloaded");

			// Capture the start time and end time on the Excel sheet
			Cell endTimeCell = currentRow.createCell(5);
			endTimeCell.setCellValue(getCurrentDateTime());
		}
	}

	@AfterClass
	public void tearDown() {
		try {
			// Save the changes to the workbook
			FileOutputStream outputStream = new FileOutputStream(System.getProperty("user.dir")+"//src//test//resources//excel//Document_download_selenium.xlsx");
			workbook.write(outputStream);
			outputStream.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

		// Quit the driver and close the browser
		driver.quit();
	}

	private void saveDocument(String formName) throws IOException {
		// Get the InputStream for the downloaded file
		InputStream inputStream = new URL(driver.getCurrentUrl()).openStream();

		// Specify the folder path to save the downloaded document
		String folderPath = System.getProperty("user.dir")+"path/to/save/folder/";

		// Create the OutputStream to save the document
		OutputStream outputStream = new FileOutputStream(folderPath + formName + ".pdf");

		// Read the bytes from the InputStream and write them to the OutputStream
		byte[] buffer = new byte[1024];
		int bytesRead;
		while ((bytesRead = inputStream.read(buffer)) != -1) {
			outputStream.write(buffer, 0, bytesRead);
		}

		// Close the streams
		inputStream.close();
		outputStream.close();
	}

	private String getCurrentDateTime() {
		LocalDateTime currentDateTime = LocalDateTime.now();
		DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
		return currentDateTime.format(formatter);
	}

	public static void downloadAndSave()  {
		try {
			// Create a Robot instance
			Robot robot = new Robot();

			// Simulate four TAB key presses
			for (int i = 0; i < 8; i++) {
				robot.keyPress(KeyEvent.VK_TAB);
				robot.keyRelease(KeyEvent.VK_TAB);
				Thread.sleep(300); // Adjust delay as needed
			}
			// Simulate an Enter key press
			for (int i = 0; i <= 10; i++) {
				robot.keyPress(KeyEvent.VK_ENTER);
				robot.keyRelease(KeyEvent.VK_ENTER);
			}
		} catch (AWTException | InterruptedException e) {
			e.printStackTrace();
		}


	}
}


