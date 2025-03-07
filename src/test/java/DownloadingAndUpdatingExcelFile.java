import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;

public class DownloadingAndUpdatingExcelFile {

	public static void main(String[] args) throws InterruptedException, IOException {

		WebDriver driver = new EdgeDriver();

		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));

		driver.get("https://rahulshettyacademy.com/upload-download-test/index.html");

		driver.manage().window().maximize();

		Thread.sleep(3000);

		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(5));

		// current price of Apple - before updating Excel

		WebElement currentAppleValue = driver
				.findElement(By.xpath("//div[@role='rowgroup']//div[2]//div[@id='cell-4-undefined']"));

		String currentApplePrice = currentAppleValue.getText();

		System.out.println("Current Apple Value Is:" + " " + currentApplePrice);

		Thread.sleep(3000);

		// download file

		WebElement downloadBtn = driver.findElement(By.id("downloadButton"));
		downloadBtn.click();

		Thread.sleep(3000);

		String filePath = "C:\\Users\\raja.bhamra\\OneDrive - Electronic Arts\\Documents\\download_latest.xlsx";
		String sheetName = "Sheet1";
		int rowIndex = 2; // Row index starts from 0
		int cellIndex = 3; // Cell index starts from 0
		String newValue = "550";

		// Load the Excel file
		FileInputStream fileInputStream = new FileInputStream(filePath);
		XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
		XSSFSheet sheet = workbook.getSheet(sheetName);

		// Read the cell value
		XSSFRow row = sheet.getRow(rowIndex);
		XSSFCell cell = row.getCell(cellIndex);

		if (cell == null) {
			cell = row.createCell(cellIndex);
		}

		// Read the current cell value
		String cellValue = cell.toString();

		// Update the cell value
		cell.setCellValue(newValue);

		// Save the changes to the Excel file
		FileOutputStream fileOutputStream = new FileOutputStream(filePath);
		workbook.write(fileOutputStream);
		fileOutputStream.close();
		workbook.close();

		System.out.println("Excel Cell value updated successfully.");

		Thread.sleep(3000);

		// upload file

		// Selenium does not have access to local machine windows like file explorer etc
		// in order to upload a specific file to a website use .sendKeys and pass in
		// file path

		WebElement uploadBtn = driver.findElement(By.xpath("//input[@type='file']"));

		uploadBtn.sendKeys("C:\\Users\\raja.bhamra\\OneDrive - Electronic Arts\\Documents\\download_latest.xlsx");

		Thread.sleep(3000);

		// edit excel

		// upload edited file

		// wait for success message to show up and wait for it to disappear

		WebElement successMsg = wait.until(
				ExpectedConditions.visibilityOfElementLocated(By.cssSelector("div[role='alert'] div:nth-child(2)")));

		String msg = successMsg.getText();

		Assert.assertEquals(msg, "Updated Excel Data Successfully.");

		System.out.println("Confirmation Message:" + " " + msg);

		Thread.sleep(3000);

		wait.until(
				ExpectedConditions.invisibilityOfElementLocated(By.cssSelector("div[role='alert'] div:nth-child(2)")));

		// verify updated excel data showing in the table

		WebElement newApplePrice = driver
				.findElement(By.xpath("//div[@role='rowgroup']//div[2]//div[@id='cell-4-undefined']"));

		String appleValue = newApplePrice.getText();

		System.out.println("New Apple Value Is:" + " " + appleValue);

		Assert.assertEquals(appleValue, "550");

		Thread.sleep(3000);

		driver.close();

	}

}
