import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

public class testing {

	public static WebDriver driver;
	public static String path = "C:\\Users\\Hi\\Desktop\\NeuralNexusLab.xlsx";
	public static FileInputStream fs;
	public static FileOutputStream fos;
	public static Workbook wb;
	public static Sheet sheet1;

	@BeforeTest
	public void setUp() throws IOException {
		driver = new ChromeDriver();
		fs = new FileInputStream(path);
		wb = new XSSFWorkbook(fs);
		sheet1 = wb.getSheetAt(0);
	}

	@Test
	public static void Test1() throws MalformedURLException, IOException {
		Row row = sheet1.getRow(1);
		Cell cell = row.getCell(0);

		if (cell != null && cell.getStringCellValue().equalsIgnoreCase("TC-01")) {

			driver.get("https://neuralnexuslab.com/");
			List<WebElement> links = driver.findElements(By.tagName("a"));
			List<String> brokenLinks = new ArrayList<>();

			for (WebElement link : links) {
				String url = link.getAttribute("href");
				String linkText = link.getText();
				if (url != null && !url.isEmpty()) {
					HttpURLConnection connection = (HttpURLConnection) new URL(url).openConnection();
					connection.setRequestMethod("HEAD");
					connection.connect();
					int responseCode = connection.getResponseCode();

					if (responseCode >= 400 && responseCode != 999) {
//						System.out.println("❌ Broken: " + linkText + " -> " + url + " - " + responseCode);
						brokenLinks.add(linkText + " (" + url + ")");
						System.out.println(brokenLinks);
					}
				}
			}

			Cell resultCell = row.createCell(7); // Column H
			Cell statusCell = row.createCell(8); // Column I
			Cell comment = row.createCell(12);

			CellStyle passStyle = wb.createCellStyle();
			passStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
			passStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			CellStyle failStyle = wb.createCellStyle();
			failStyle.setFillForegroundColor(IndexedColors.PINK1.getIndex());
			failStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			if (brokenLinks.isEmpty()) {
				resultCell.setCellValue("No Broken Links");
				statusCell.setCellValue("Passed");
				statusCell.setCellStyle(passStyle);
				comment.setCellValue("All Good");

				// Apply LIGHT_GREEN style to all cells in the row
				for (int i = 0; i < row.getLastCellNum(); i++) {
					Cell c = row.getCell(i);
					if (c == null) {
						c = row.createCell(i);
					}
					c.setCellStyle(passStyle);
				}
			} else {
				resultCell.setCellValue("Broken Links are there");
				statusCell.setCellValue("Failed");
				statusCell.setCellStyle(failStyle);
				comment.setCellValue(String.join("\n", brokenLinks));
				for (int i = 0; i < row.getLastCellNum(); i++) {
					Cell c = row.getCell(i);
					if (c == null) {
						c = row.createCell(i);
					}
					c.setCellStyle(failStyle);
				}
			}

		} else {
			System.out.println("Failed");
		}
	}

	@Test
	public static void Test2() throws MalformedURLException, IOException {
		Row row = sheet1.getRow(2);
		Cell cell = row.getCell(0);

		if (cell != null && cell.getStringCellValue().equalsIgnoreCase("TC-02")) {

			driver.get("https://neuralnexuslab.com/pages/contact");
			List<WebElement> links = driver.findElements(By.tagName("a"));
			List<String> brokenLinks = new ArrayList<>();

			for (WebElement link : links) {
				String url = link.getAttribute("href");
				String linkText = link.getText().isEmpty() ? "Unnamed Link" : link.getText();

				// ✅ Skip non-http(s) links like mailto:, javascript:, tel:, etc.
				if (url == null || url.isEmpty() || !(url.startsWith("http://") || url.startsWith("https://"))) {
					System.out.println("⚠️ Skipped non-HTTP URL: " + url);
					continue;
				}

				try {
					HttpURLConnection connection = (HttpURLConnection) new URL(url).openConnection();
					connection.setRequestMethod("HEAD");
					connection.setConnectTimeout(3000);
					connection.connect();
					int responseCode = connection.getResponseCode();

					if (responseCode >= 400 && responseCode != 999) {
						brokenLinks.add(linkText + " (" + url + ")");
						System.out.println("❌ Broken: " + linkText + " -> " + url + " - " + responseCode);
					}
				} catch (Exception e) {
					brokenLinks.add(linkText + " (" + url + ") - Exception: " + e.getClass().getSimpleName());
					System.out.println("❌ Exception: " + linkText + " -> " + url + " - " + e.getMessage());
				}
			}

			Cell resultCell = row.createCell(7); // Column H
			Cell statusCell = row.createCell(8); // Column I
			Cell comment = row.createCell(12);

			CellStyle passStyle = wb.createCellStyle();
			passStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
			passStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			CellStyle failStyle = wb.createCellStyle();
			failStyle.setFillForegroundColor(IndexedColors.PINK1.getIndex());
			failStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			if (brokenLinks.isEmpty()) {
				resultCell.setCellValue("No Broken Links");
				statusCell.setCellValue("Passed");
				statusCell.setCellStyle(passStyle);
				comment.setCellValue("All Good");

				for (int i = 0; i < row.getLastCellNum(); i++) {
					Cell c = row.getCell(i);
					if (c == null) {
						c = row.createCell(i);
					}
					c.setCellStyle(passStyle);
				}
			} else {
				resultCell.setCellValue("Broken Links are there");
				statusCell.setCellValue("Failed");
				statusCell.setCellStyle(failStyle);
				comment.setCellValue(String.join("\n", brokenLinks));

				for (int i = 0; i < row.getLastCellNum(); i++) {
					Cell c = row.getCell(i);
					if (c == null) {
						c = row.createCell(i);
					}
					c.setCellStyle(failStyle);
				}
			}
		} else {
			System.out.println("Failed");
		}
	}

	@Test
	public static void Test3() throws MalformedURLException, IOException {
		Row row = sheet1.getRow(3);
		Cell cell = row.getCell(0);

		if (cell != null && cell.getStringCellValue().equalsIgnoreCase("TC-03")) {

			driver.get("https://neuralnexuslab.com/about");
			List<WebElement> links = driver.findElements(By.tagName("a"));
			List<String> brokenLinks = new ArrayList<>();

			for (WebElement link : links) {
				String url = link.getAttribute("href");
				String linkText = link.getText();
				if (url != null && !url.isEmpty()) {
					HttpURLConnection connection = (HttpURLConnection) new URL(url).openConnection();
					connection.setRequestMethod("HEAD");
					connection.connect();
					int responseCode = connection.getResponseCode();

					if (responseCode >= 400 && responseCode != 999) {
//						System.out.println("❌ Broken: " + linkText + " -> " + url + " - " + responseCode);
						brokenLinks.add(linkText + " (" + url + ")");
						System.out.println(brokenLinks);
					}
				}
			}

			Cell resultCell = row.createCell(7); // Column H
			Cell statusCell = row.createCell(8); // Column I
			Cell comment = row.createCell(12);

			CellStyle passStyle = wb.createCellStyle();
			passStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
			passStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			CellStyle failStyle = wb.createCellStyle();
			failStyle.setFillForegroundColor(IndexedColors.PINK1.getIndex());
			failStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			if (brokenLinks.isEmpty()) {
				resultCell.setCellValue("No Broken Links");
				statusCell.setCellValue("Passed");
				statusCell.setCellStyle(passStyle);
				comment.setCellValue("All Good");

				// Apply LIGHT_GREEN style to all cells in the row
				for (int i = 0; i < row.getLastCellNum(); i++) {
					Cell c = row.getCell(i);
					if (c == null) {
						c = row.createCell(i);
					}
					c.setCellStyle(passStyle);
				}
			} else {
				resultCell.setCellValue("Broken Links are there");
				statusCell.setCellValue("Failed");
				statusCell.setCellStyle(failStyle);
				comment.setCellValue(String.join("\n", brokenLinks));
				for (int i = 0; i < row.getLastCellNum(); i++) {
					Cell c = row.getCell(i);
					if (c == null) {
						c = row.createCell(i);
					}
					c.setCellStyle(failStyle);
				}
			}

		} else {
			System.out.println("Failed");
		}
	}

	@Test
	public static void Test4() throws MalformedURLException, IOException {
		Row row = sheet1.getRow(4);
		Cell cell = row.getCell(0);

		if (cell != null && cell.getStringCellValue().equalsIgnoreCase("TC-04")) {

			driver.get("https://neuralnexuslab.com/privacy-policy");
			List<WebElement> links = driver.findElements(By.tagName("a"));
			List<String> brokenLinks = new ArrayList<>();

			for (WebElement link : links) {
				String url = link.getAttribute("href");
				String linkText = link.getText();
				if (url != null && !url.isEmpty()) {
					HttpURLConnection connection = (HttpURLConnection) new URL(url).openConnection();
					connection.setRequestMethod("HEAD");
					connection.connect();
					int responseCode = connection.getResponseCode();

					if (responseCode >= 400 && responseCode != 999) {
//						System.out.println("❌ Broken: " + linkText + " -> " + url + " - " + responseCode);
						brokenLinks.add(linkText + " (" + url + ")");
						System.out.println(brokenLinks);
					}
				}
			}

			Cell resultCell = row.createCell(7); // Column H
			Cell statusCell = row.createCell(8); // Column I
			Cell comment = row.createCell(12);

			CellStyle passStyle = wb.createCellStyle();
			passStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
			passStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			CellStyle failStyle = wb.createCellStyle();
			failStyle.setFillForegroundColor(IndexedColors.PINK1.getIndex());
			failStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			if (brokenLinks.isEmpty()) {
				resultCell.setCellValue("No Broken Links");
				statusCell.setCellValue("Passed");
				statusCell.setCellStyle(passStyle);
				comment.setCellValue("All Good");

				// Apply LIGHT_GREEN style to all cells in the row
				for (int i = 0; i < row.getLastCellNum(); i++) {
					Cell c = row.getCell(i);
					if (c == null) {
						c = row.createCell(i);
					}
					c.setCellStyle(passStyle);
				}
			} else {
				resultCell.setCellValue("Broken Links are there");
				statusCell.setCellValue("Failed");
				statusCell.setCellStyle(failStyle);
				comment.setCellValue(String.join("\n", brokenLinks));
				for (int i = 0; i < row.getLastCellNum(); i++) {
					Cell c = row.getCell(i);
					if (c == null) {
						c = row.createCell(i);
					}
					c.setCellStyle(failStyle);
				}
			}

		} else {
			System.out.println("Failed");
		}
	}

	@Test
	public static void Test5() throws MalformedURLException, IOException {
		Row row = sheet1.getRow(5);
		Cell cell = row.getCell(0);

		if (cell != null && cell.getStringCellValue().equalsIgnoreCase("TC-05")) {

			driver.get("https://neuralnexuslab.com/tnc");
			List<WebElement> links = driver.findElements(By.tagName("a"));
			List<String> brokenLinks = new ArrayList<>();

			for (WebElement link : links) {
				String url = link.getAttribute("href");
				String linkText = link.getText();
				if (url != null && !url.isEmpty()) {
					HttpURLConnection connection = (HttpURLConnection) new URL(url).openConnection();
					connection.setRequestMethod("HEAD");
					connection.connect();
					int responseCode = connection.getResponseCode();

					if (responseCode >= 400 && responseCode != 999) {
//						System.out.println("❌ Broken: " + linkText + " -> " + url + " - " + responseCode);
						brokenLinks.add(linkText + " (" + url + ")");
						System.out.println(brokenLinks);
					}
				}
			}

			Cell resultCell = row.createCell(7); // Column H
			Cell statusCell = row.createCell(8); // Column I
			Cell comment = row.createCell(12);

			CellStyle passStyle = wb.createCellStyle();
			passStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
			passStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			CellStyle failStyle = wb.createCellStyle();
			failStyle.setFillForegroundColor(IndexedColors.PINK1.getIndex());
			failStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			if (brokenLinks.isEmpty()) {
				resultCell.setCellValue("No Broken Links");
				statusCell.setCellValue("Passed");
				statusCell.setCellStyle(passStyle);
				comment.setCellValue("All Good");

				// Apply LIGHT_GREEN style to all cells in the row
				for (int i = 0; i < row.getLastCellNum(); i++) {
					Cell c = row.getCell(i);
					if (c == null) {
						c = row.createCell(i);
					}
					c.setCellStyle(passStyle);
				}
			} else {
				resultCell.setCellValue("Broken Links are there");
				statusCell.setCellValue("Failed");
				statusCell.setCellStyle(failStyle);
				comment.setCellValue(String.join("\n", brokenLinks));
				for (int i = 0; i < row.getLastCellNum(); i++) {
					Cell c = row.getCell(i);
					if (c == null) {
						c = row.createCell(i);
					}
					c.setCellStyle(failStyle);
				}
			}

		} else {
			System.out.println("Failed");
		}
	}

	@Test
	public static void Test6() throws MalformedURLException, IOException {
		Row row = sheet1.getRow(6); // Row 7 (index 6)
		Cell cell = row.getCell(0);

		if (cell != null && cell.getStringCellValue().equalsIgnoreCase("TC-06")) {

			driver.get("https://neuralnexuslab.blog");
			List<WebElement> links = driver.findElements(By.tagName("a"));
			List<String> brokenLinks = new ArrayList<>();

			for (WebElement link : links) {
				String url = link.getAttribute("href");
				String linkText = link.getText().isEmpty() ? "Unnamed Link" : link.getText();

				if (url == null || url.isEmpty() || !url.startsWith("http")) {
					continue;
				}

				boolean isBroken = false;

				try {
					HttpURLConnection connection = (HttpURLConnection) new URL(url).openConnection();
					connection.setRequestMethod("HEAD");
					connection.setConnectTimeout(5000); // avoid hang
					connection.setReadTimeout(5000);
					connection.connect();

					int responseCode = connection.getResponseCode();

					if (responseCode >= 400 && responseCode != 999) {
						isBroken = true;
//			            System.out.println("❌ Broken: " + linkText + " -> " + url + " - " + responseCode);
					} else {
//			            System.out.println("✅ Valid: " + linkText + " -> " + url);
					}

				} catch (Exception e) {
					isBroken = true;
//			        System.out.println("❌ Exception: " + linkText + " -> " + url + " - " + e.getClass().getSimpleName());
				}

				if (isBroken) {
					brokenLinks.add(linkText + " (" + url + ")");
				}
			}

			// Create/overwrite cells
			Cell resultCell = row.createCell(7); // Column H
			Cell statusCell = row.createCell(8); // Column I
			Cell comment = row.createCell(12); // Column M

			// Styles
			CellStyle passStyle = wb.createCellStyle();
			passStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
			passStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			CellStyle failStyle = wb.createCellStyle();
			failStyle.setFillForegroundColor(IndexedColors.PINK1.getIndex());
			failStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			if (brokenLinks.isEmpty()) {
				resultCell.setCellValue("No Broken Links");
				statusCell.setCellValue("Passed");
				statusCell.setCellStyle(passStyle);
				comment.setCellValue("All Good");

				// Apply LIGHT_GREEN style to all cells in the row
				for (int i = 0; i < row.getLastCellNum(); i++) {
					Cell c = row.getCell(i);
					if (c == null) {
						c = row.createCell(i);
					}
					c.setCellStyle(passStyle);
				}
			} else {
				resultCell.setCellValue("Broken Links are there");
				statusCell.setCellValue("Failed");
				statusCell.setCellStyle(failStyle);
				comment.setCellValue(String.join("\n", brokenLinks));
				for (int i = 0; i < row.getLastCellNum(); i++) {
					Cell c = row.getCell(i);
					if (c == null) {
						c = row.createCell(i);
					}
					c.setCellStyle(failStyle);
				}
			}

		} else {
			System.out.println("Failed");
		}
	}

	@Test
	public static void Test7() throws MalformedURLException, IOException {
		Row row = sheet1.getRow(7);
		Cell cell = row.getCell(0);

		if (cell != null && cell.getStringCellValue().equalsIgnoreCase("TC-07")) {

			driver.get("https://neuralnexuslab.tech");
			List<WebElement> links = driver.findElements(By.tagName("a"));
			List<String> brokenLinks = new ArrayList<>();

			for (WebElement link : links) {
				String url = link.getAttribute("href");
				String linkText = link.getText();
				if (url != null && !url.isEmpty()) {
					HttpURLConnection connection = (HttpURLConnection) new URL(url).openConnection();
					connection.setRequestMethod("HEAD");
					connection.connect();
					int responseCode = connection.getResponseCode();

					if (responseCode >= 400 && responseCode != 999) {
//						System.out.println("❌ Broken: " + linkText + " -> " + url + " - " + responseCode);
						brokenLinks.add(linkText + " (" + url + ")");
						System.out.println(brokenLinks);
					}
				}
			}

			Cell resultCell = row.createCell(7); // Column H
			Cell statusCell = row.createCell(8); // Column I
			Cell comment = row.createCell(12);

			CellStyle passStyle = wb.createCellStyle();
			passStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
			passStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			CellStyle failStyle = wb.createCellStyle();
			failStyle.setFillForegroundColor(IndexedColors.PINK1.getIndex());
			failStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			if (brokenLinks.isEmpty()) {
				resultCell.setCellValue("No Broken Links");
				statusCell.setCellValue("Passed");
				statusCell.setCellStyle(passStyle);
				comment.setCellValue("All Good");

				// Apply LIGHT_GREEN style to all cells in the row
				for (int i = 0; i < row.getLastCellNum(); i++) {
					Cell c = row.getCell(i);
					if (c == null) {
						c = row.createCell(i);
					}
					c.setCellStyle(passStyle);
				}
			} else {
				resultCell.setCellValue("Broken Links are there");
				statusCell.setCellValue("Failed");
				statusCell.setCellStyle(failStyle);
				comment.setCellValue(String.join("\n", brokenLinks));
				for (int i = 0; i < row.getLastCellNum(); i++) {
					Cell c = row.getCell(i);
					if (c == null) {
						c = row.createCell(i);
					}
					c.setCellStyle(failStyle);
				}
			}

		} else {
			System.out.println("Failed");
		}
	}

	@Test
	public static void Test8() throws MalformedURLException, IOException {
		Row row = sheet1.getRow(8);
		Cell cell = row.getCell(0);

		if (cell != null && cell.getStringCellValue().equalsIgnoreCase("TC-08")) {

			driver.get("https://neuralnexuslab.digital");
			List<WebElement> links = driver.findElements(By.tagName("a"));
			List<String> brokenLinks = new ArrayList<>();

			for (WebElement link : links) {
				String url = link.getAttribute("href");
				String linkText = link.getText();
				if (url != null && !url.isEmpty()) {
					HttpURLConnection connection = (HttpURLConnection) new URL(url).openConnection();
					connection.setRequestMethod("HEAD");
					connection.connect();
					int responseCode = connection.getResponseCode();

					if (responseCode >= 400 && responseCode != 999) {
//						System.out.println("❌ Broken: " + linkText + " -> " + url + " - " + responseCode);
						brokenLinks.add(linkText + " (" + url + ")");
						System.out.println(brokenLinks);
					}
				}
			}

			Cell resultCell = row.createCell(7); // Column H
			Cell statusCell = row.createCell(8); // Column I
			Cell comment = row.createCell(12);

			CellStyle passStyle = wb.createCellStyle();
			passStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
			passStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			CellStyle failStyle = wb.createCellStyle();
			failStyle.setFillForegroundColor(IndexedColors.PINK1.getIndex());
			failStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			if (brokenLinks.isEmpty()) {
				resultCell.setCellValue("No Broken Links");
				statusCell.setCellValue("Passed");
				statusCell.setCellStyle(passStyle);
				comment.setCellValue("All Good");

				// Apply LIGHT_GREEN style to all cells in the row
				for (int i = 0; i < row.getLastCellNum(); i++) {
					Cell c = row.getCell(i);
					if (c == null) {
						c = row.createCell(i);
					}
					c.setCellStyle(passStyle);
				}
			} else {
				resultCell.setCellValue("Broken Links are there");
				statusCell.setCellValue("Failed");
				statusCell.setCellStyle(failStyle);
				comment.setCellValue(String.join("\n", brokenLinks));
				for (int i = 0; i < row.getLastCellNum(); i++) {
					Cell c = row.getCell(i);
					if (c == null) {
						c = row.createCell(i);
					}
					c.setCellStyle(failStyle);
				}
			}

		} else {
			System.out.println("Failed");
		}
	}

	@Test
	public static void Test9() throws MalformedURLException, IOException {
		Row row = sheet1.getRow(9);
		Cell cell = row.getCell(0);

		if (cell != null && cell.getStringCellValue().equalsIgnoreCase("TC-09")) {

			driver.get("https://ai.neuralnexuslab.com");
			List<WebElement> links = driver.findElements(By.tagName("a"));
			List<String> brokenLinks = new ArrayList<>();

			for (WebElement link : links) {
				String url = link.getAttribute("href");
				String linkText = link.getText();
				if (url != null && !url.isEmpty()) {
					HttpURLConnection connection = (HttpURLConnection) new URL(url).openConnection();
					connection.setRequestMethod("HEAD");
					connection.connect();
					int responseCode = connection.getResponseCode();

					if (responseCode >= 400 && responseCode != 999) {
//						System.out.println("❌ Broken: " + linkText + " -> " + url + " - " + responseCode);
						brokenLinks.add(linkText + " (" + url + ")");
						System.out.println(brokenLinks);
					}
				}
			}

			Cell resultCell = row.createCell(7); // Column H
			Cell statusCell = row.createCell(8); // Column I
			Cell comment = row.createCell(12);

			CellStyle passStyle = wb.createCellStyle();
			passStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
			passStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			CellStyle failStyle = wb.createCellStyle();
			failStyle.setFillForegroundColor(IndexedColors.PINK1.getIndex());
			failStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			if (brokenLinks.isEmpty()) {
				resultCell.setCellValue("No Broken Links");
				statusCell.setCellValue("Passed");
				statusCell.setCellStyle(passStyle);
				comment.setCellValue("All Good");

				// Apply LIGHT_GREEN style to all cells in the row
				for (int i = 0; i < row.getLastCellNum(); i++) {
					Cell c = row.getCell(i);
					if (c == null) {
						c = row.createCell(i);
					}
					c.setCellStyle(passStyle);
				}
			} else {
				resultCell.setCellValue("Broken Links are there");
				statusCell.setCellValue("Failed");
				statusCell.setCellStyle(failStyle);
				comment.setCellValue(String.join("\n", brokenLinks));
				for (int i = 0; i < row.getLastCellNum(); i++) {
					Cell c = row.getCell(i);
					if (c == null) {
						c = row.createCell(i);
					}
					c.setCellStyle(failStyle);
				}
			}

		} else {
			System.out.println("Failed");
		}
	}

	@AfterTest
	public void tearDown() {
		try {
			// Write updates to Excel
			if (wb != null) {
				fos = new FileOutputStream(path);
				wb.write(fos);
				fos.close();
			}

			// Close browser
			if (driver != null) {
				driver.quit();
			}

			// Close workbook and input stream
			if (wb != null) {
				wb.close();
			}
			if (fs != null) {
				fs.close();
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
