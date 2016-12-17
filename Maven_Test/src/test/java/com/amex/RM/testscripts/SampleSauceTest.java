package com.amex.RM.testscripts;

import org.testng.Assert;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.text.DateFormat;
import java.text.Format;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import javax.swing.JDialog;
import javax.swing.JOptionPane;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
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

		Prereport();
		DesiredCapabilities caps = DesiredCapabilities.chrome();
		caps.setCapability("platform", "Windows 7");
		caps.setCapability("version", "43.0");

		WebDriver driver = new RemoteWebDriver(new URL(URL), caps);

		/**
		 * Goes to Sauce Lab's guinea-pig page and prints title
		 */

		driver.get("https://saucelabs.com/test/guinea-pig");
		System.out.println("title of page is: " + driver.getTitle());
		write_intoExcel(0, driver.getTitle());

		DateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
		Date date = new Date();

		write_intoExcel(1, dateFormat.format(date));

		takescreenshot(driver);
		Assert.assertEquals(false, true);

		driver.quit();
	}

	public static void write_intoExcel(int rowno, String values) throws Exception {
		try {
			String filepath = System.getProperty("user.dir") + "//Results//Execution_Result.xlsm";
			FileInputStream file1 = new FileInputStream(new File(filepath));
			XSSFWorkbook wb = new XSSFWorkbook(file1);
			FileOutputStream outFile1 = new FileOutputStream(new File(filepath));
			XSSFSheet sheet = wb.getSheetAt(0);
			Row row;
			row = sheet.createRow(rowno);
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

	public static void Prereport() throws Exception {

		String path = System.getProperty("user.dir");
		File theDir = new File(path + "/Results");
		if (!theDir.exists()) {
			theDir.mkdir();
		}
		// Creating the Screenshot folder inside the Workspace
		theDir = new File(path + "/Results/Screenshots");
		if (!theDir.exists()) {
			theDir.mkdir();
		}

		theDir = new File(path + "/Results/Last Execution results");
		if (!theDir.exists()) {
			theDir.mkdir();
		}

		String ts;
		ts = new SimpleDateFormat("dd-MM-yyyy_HH:mm:ss").format(Calendar.getInstance().getTime());

		// Renaming the Last executed Excel and PDF file with Current Times
		final File file = new File(path + "/Results/Execution_Result.xlsm");
		String time1 = ts.toString();
		String time2 = time1.replace(":", "_");
		String time3 = time2.replace(".", "_");
		String time4 = time3.replace("-", "_");
		final String st = path + "/Results/Last Execution results/Excel/Execution_Result_" + time4 + ".xlsm";
		file.renameTo(new File(st));

		// Copying the template xlsx file from Template folder to Result folder
		File source = new File(path + "/Template");
		File desc = new File(path + "/Results");
		try {
			FileUtils.copyDirectory(source, desc);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			if (e.toString().toLowerCase().contains("filenotfoundexception")) {
				JOptionPane optionPane = new JOptionPane(
						"Please close the Result.xlsx file and try again. Or Check the Path in Report file.",
						JOptionPane.WARNING_MESSAGE);
				JDialog dialog = optionPane.createDialog("Alert!");
				dialog.setAlwaysOnTop(true); // to show top of all other
												// application
				dialog.setVisible(true); // to visible the dialog
				System.exit(0);
			}
		}
	}

	public void takescreenshot(WebDriver driver) throws Exception {

		String path = System.getProperty("user.dir");
		Date date = new Date();
		Format formatter = new SimpleDateFormat("yyyy-MM-dd_hh-mm-ss");
		String SSpath = "";
		File scrnsht = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
		File screenshotpath = new File(path + "/Results/Screenshots" + "//FailTC_" + formatter.format(date) + ".jpeg");
		java.net.URI u = screenshotpath.toURI();
		SSpath = u.toString();
		try {
			FileUtils.copyFile(scrnsht, screenshotpath);
		} catch (IOException e) {
			e.printStackTrace();
		}

		String filepath = System.getProperty("user.dir") + "//Results//Execution_Result.xlsm";
		FileInputStream file1 = new FileInputStream(new File(filepath));
		XSSFWorkbook wb = new XSSFWorkbook(file1);
		FileOutputStream outFile1 = new FileOutputStream(new File(filepath));
		XSSFSheet sheet = wb.getSheetAt(0);
		Row row;
		row = sheet.createRow(2);

		Cell cell = row.createCell(0);
		// Style the cell with borders all around.
		CellStyle style1 = wb.createCellStyle();
		XSSFFont style_font = wb.createFont();
		style_font.setUnderline(XSSFFont.U_SINGLE);
		style_font.setColor(IndexedColors.BLUE.getIndex());
		style1.setFont(style_font);
		// URL
		CreationHelper createHelper = wb.getCreationHelper();
		XSSFHyperlink link = (XSSFHyperlink) createHelper.createHyperlink(Hyperlink.LINK_URL);
		cell.setCellValue("Screenshot");
		link.setAddress(SSpath);
		cell.setHyperlink((XSSFHyperlink) link);
		// style1 the cell with borders all around.
		style1.setBorderBottom(CellStyle.BORDER_THIN);
		style1.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		style1.setBorderLeft(CellStyle.BORDER_THIN);
		style1.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		style1.setBorderRight(CellStyle.BORDER_THIN);
		style1.setRightBorderColor(IndexedColors.BLACK.getIndex());
		style1.setBorderTop(CellStyle.BORDER_THIN);
		style1.setTopBorderColor(IndexedColors.BLACK.getIndex());
		style1.setWrapText(true);
		cell.setCellStyle(style1);

		wb.write(outFile1);
		outFile1.close();
	}

}