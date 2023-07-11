package com.test;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;
import static com.utility.ScreenRecorderUtil.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.WindowType;
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
		wait = new WebDriverWait(driver, Duration.ofSeconds(30));
		try {
			String filePath=System.getProperty("user.dir")+"/src/test/resources/excel/Document_download_selenium.xlsx";
			File file =    new File(filePath);
			FileInputStream inputStream = new FileInputStream(file); 			
			// Load the Excel workbook
			workbook=new XSSFWorkbook(inputStream);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	@Test
	public void testFormDownload() throws Exception {
		startRecord("testFormDownload");
		// Open the link
		driver.get("http://houseofforms.org/ActWiseForms.aspx?Search=&ActID=9650A6D7E5A396D0");
		// Get to the first sheet
		Sheet sheet = workbook.getSheetAt(0);
		// Get the total number of rows with data in the sheet			
		int totalRows = sheet.getLastRowNum() + 1;

		// Iterate through each row
		for (int row = 1; row < totalRows; row++) {
			// Get the form name from the current row
			Row currentRow = sheet.getRow(row);
			Cell formNameCell = currentRow.getCell(1);
			String formName = formNameCell.getStringCellValue().replaceAll("[^\\x20-\\x7E]", "").trim();
			// Search the form name on the website's search section
			WebElement searchInput = wait.until(ExpectedConditions.presenceOfElementLocated(By.id("ctl00_txtSearch")));
			searchInput.clear();
			searchInput.sendKeys(formName);
			WebElement searchButton = driver.findElement(By.id("ctl00_Search"));
			searchButton.click();
			// Capture the start time
			Cell startTimeCell = currentRow.createCell(4);
			String startTime = getCurrentDateTime();
			startTimeCell.setCellValue(getCurrentDateTime());			
			//  click on each document link
			List<WebElement> viewElements = driver.findElements(By.xpath("//a[text()=' View ']"));
			for (WebElement documentLink : viewElements) {				
				documentLink.click();							
				Thread.sleep(50000);
				// Download the form	
				downloadAndSave();				
				// Go back to the previous page for the next document
				driver.navigate().back();
				Thread.sleep(3000);	
			}
			// Update the status on the Excel sheet
			Cell statusCell = currentRow.createCell(3);
			statusCell.setCellValue("Downloaded");
			// Capturing the start time and end time on the Excel sheet
			Cell endTimeCell = currentRow.createCell(5);
			String endTime = getCurrentDateTime();
			endTimeCell.setCellValue(getCurrentDateTime());
			String timetaken=calculateTimeTaken(startTime, endTime);
			System.out.println(timetaken);
			//Updating the total time taken column
			Cell timeTakenCell = currentRow.createCell(6);
			timeTakenCell.setCellValue(timetaken);
		}
	}

	@AfterClass
	public void tearDown() throws Exception {
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
		stopRecord();
	}
	public static String getCurrentDateTime() {
		LocalDateTime currentDateTime = LocalDateTime.now();
		DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
		return currentDateTime.format(formatter);
	}

	public static String calculateTimeTaken(String startTime, String endTime) {
		DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
		LocalDateTime startDateTime = LocalDateTime.parse(startTime, formatter);
		LocalDateTime endDateTime = LocalDateTime.parse(endTime, formatter);
		Duration duration = Duration.between(startDateTime, endDateTime);
		long seconds = duration.getSeconds();
		long hours = seconds / 3600;
		long minutes = (seconds % 3600) / 60;
		long remainingSeconds = seconds % 60;
		String timeTaken = String.format("%02d:%02d:%02d", hours, minutes, remainingSeconds);
		return timeTaken;
	}
	public void downloadAndSave()  {
		try {
			// Create a Robot instance
			Robot robot = new Robot();
			// Simulate TAB key presses
			for (int i = 0; i < 8; i++) {
				robot.keyPress(KeyEvent.VK_TAB);
				robot.keyRelease(KeyEvent.VK_TAB);
				Thread.sleep(3000); // Adjust delay as needed
			}
			// Simulate Enter key press			
				robot.keyPress(KeyEvent.VK_ENTER);
				robot.keyRelease(KeyEvent.VK_ENTER);
				Thread.sleep(10000);
				robot.keyPress(KeyEvent.VK_ENTER);
				Thread.sleep(10000);
		} catch (AWTException | InterruptedException e) {
			e.printStackTrace();
		}
	}
}


