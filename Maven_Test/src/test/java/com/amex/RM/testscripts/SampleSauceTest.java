package com.amex.RM.testscripts;

import org.testng.annotations.Test;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.net.URL;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.testng.annotations.Test;


 
public class SampleSauceTest {
 
  public static final String USERNAME = "naseemamex6270";
  public static final String ACCESS_KEY = "b8d0327b-ba50-4558-a42b-537dfd27456f";
  public static final String URL = "https://" + USERNAME + ":" + ACCESS_KEY + "@ondemand.saucelabs.com:443/wd/hub";
 
  @Test
  public void testjenkins() throws Exception {
 
    DesiredCapabilities caps = DesiredCapabilities.chrome();
    caps.setCapability("platform", "Windows 7");
    caps.setCapability("version", "43.0");
 
    WebDriver driver = new RemoteWebDriver(new URL(URL), caps);
 
    /**
     * Goes to Sauce Lab's guinea-pig page and prints title
     */
 
    driver.get("https://saucelabs.com/test/guinea-pig");
    System.out.println("title of page is: " + driver.getTitle());
    write_intoExcel(driver.getTitle());
 
    driver.quit();
  }
  
  public static void write_intoExcel(String values)
			throws Exception {
		try {
			String filepath = System.getProperty("user.dir")+"//Results//Data.xlsx";
			FileInputStream file1 = new FileInputStream(new File(filepath));
			XSSFWorkbook wb = new XSSFWorkbook(file1);
			FileOutputStream outFile1 = new FileOutputStream(new File(filepath));
			XSSFSheet sheet = wb.getSheetAt(0);
			Row row;
			row = sheet.createRow(0);
			// Style the cell with borders all around.
			CellStyle style = wb.createCellStyle();
			style.setBorderBottom(CellStyle.BORDER_THIN);
			style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
			style.setBorderLeft(CellStyle.BORDER_THIN);
			style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
			style.setBorderRight(CellStyle.BORDER_THIN);
			style.setRightBorderColor(IndexedColors.BLACK.getIndex());
			style.setBorderTop(CellStyle.BORDER_THIN);
			style.setTopBorderColor(IndexedColors.BLACK.getIndex());
			style.setWrapText(true);
			Cell cell = row.createCell(0);
			cell.setCellValue(values);
			cell.setCellStyle(style);
			wb.write(outFile1);
			outFile1.close();
		} catch (FileNotFoundException e) {
			System.out.println(
					"Unable to write into the excel file. Please check whether the file is currently opened. Close the file and try again");
			System.exit(0);
		}
	}
}