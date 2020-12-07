package com.prac.keyword.engine;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.format.CellFormatType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.Select;

import com.prac.keyword.base.Base;

/**
 * 
 * @author Umarani Sridhar
 *
 */
public class KeyWordEngine {

	public WebDriver driver;
	public Properties prop;

	public static Workbook book;
	public static Workbook mainBook;
	public static Workbook mainOutBook;
	public static Sheet sheet;
	public static Sheet mainSheet;

	public Base base;
	public WebElement element;

	public final String EXECUTION_SHEET_PATH = "./src/main/resources/Execution.xlsx"; // Master sheet where run flag is
	public final String EXECUTION_OUT_SHEET_PATH = "./src/main/resources/Execution_Output.xlsx";
	// set
	public final String SCENARIO_SHEET_PATH = "./src/main/java/com/prac/keyword/scenarios/";
	StringBuffer strQuote = new StringBuffer(); // defining a global Quote string variable to store the Quotes

	public FileInputStream fileInputStreamReader(String filepath) {

		try {
			return new FileInputStream(filepath);

		} catch (FileNotFoundException e) {
			e.printStackTrace();
			return null;
		}

	}

	public void excelFileOutputStreamWriter(String filepath, Workbook wb) {

		try {
			FileOutputStream fileout = new FileOutputStream(filepath);
			wb.write(fileout);
			fileout.close();

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	public void mainExecution() {

		FileInputStream file = fileInputStreamReader(EXECUTION_SHEET_PATH);

		try {
			mainBook = WorkbookFactory.create(file);
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		mainSheet = mainBook.getSheetAt(0); // Get the first sheet
		int kMain = 0;
		int lastRowMain = mainSheet.getLastRowNum() - 1;

		for (int iMain = 0; iMain <= lastRowMain; iMain++) {
			try {

				String testRunFlag = mainSheet.getRow(iMain + 1).getCell(kMain).toString().trim();
				String testScenario = mainSheet.getRow(iMain + 1).getCell(kMain + 1).toString().trim();
				String iteration = mainSheet.getRow(iMain + 1).getCell(kMain + 2).toString().trim();
				int iterationCount = Integer.parseInt(iteration);

				if (testRunFlag.equalsIgnoreCase("Y")) {

					for (int jMain = 1; jMain <= iterationCount; jMain++) {

						startExecution(testScenario, jMain);

					}
					try {
						file.close();
						mainSheet.getRow(iMain+1).createCell(kMain + 3).setCellValue(strQuote.toString());
						excelFileOutputStreamWriter(EXECUTION_OUT_SHEET_PATH, mainBook);


					} catch (FileNotFoundException e) {
						e.printStackTrace();
					}
				}

			} catch (Exception e) {
				System.out.println(" Exception: " + e.toString());
			}
		}

	}
	
	public String getDataFromTestData(String sheetName,int iterator) {
		return sheetName;
		
	}

	public void startExecution(String sheetName, int iterator) {

		FileInputStream file = null;
		try {
			file = new FileInputStream(SCENARIO_SHEET_PATH + sheetName + ".xlsx");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}

		try {
			book = WorkbookFactory.create(file);
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		sheet = book.getSheet(sheetName);
		FormulaEvaluator evaluator = book.getCreationHelper().createFormulaEvaluator(); 

		int scenarioLastRow = sheet.getLastRowNum();

		int k = 0;
		for (int i = 0; i < scenarioLastRow; i++) {
			try {

				String testStep = sheet.getRow(i + 1).getCell(k).toString().trim();
				System.out.println(testStep);
				String locatorType = sheet.getRow(i + 1).getCell(k + 1).toString().trim();
				String locatorValue = sheet.getRow(i + 1).getCell(k + 2).toString().trim();
				String action = sheet.getRow(i + 1).getCell(k + 3).toString().trim();
				Cell testdata = sheet.getRow(i + 1).getCell(k + 4);
				String value;
				switch (evaluator.evaluateFormulaCell(testdata)) {
			        case Cell.CELL_TYPE_BOOLEAN:
			            Boolean blnvalue = testdata.getBooleanCellValue();
			            value = blnvalue.toString();
			            break;
			        case Cell.CELL_TYPE_NUMERIC:
			            int value2 = (int)testdata.getNumericCellValue();
			            value = Integer.toString(value2);
			            break;
			        case Cell.CELL_TYPE_STRING:
			        	value = testdata.getStringCellValue();
			            break;
			        default:
			        	String initialvalue = testdata.toString().trim();
			        	if (value.contains("#"))
			        		value = getDataFromTestData(sheetName,iterator); // define this function to fetch data from TestData.xlsx
			        	else
			        		value = initialvalue;
			            
			    }

				

				switch (action.toLowerCase()) {
				case "open browser":
					base = new Base();
					prop = base.init_properties();
					if (value.isEmpty() || value.equals("NA")) {
						driver = base.init_driver(prop.getProperty("browser"));
					} else {
						driver = base.init_driver(value);
					}
					break;

				case "enter url":
					if (value.isEmpty() || value.equals("NA")) {
						driver.get(prop.getProperty("url"));
					} else {
						driver.get(value);
					}
					break;

				case "quit":
					driver.quit();
					break;

				default:
					break;
				}

				driver.manage().timeouts().implicitlyWait(200, TimeUnit.SECONDS);
				JavascriptExecutor executor = (JavascriptExecutor) driver;

				switch (locatorType) {
				case "id":
					element = driver.findElement(By.id(locatorValue));
					if (action.equalsIgnoreCase("sendkeys")) {
						element.clear();
						element.sendKeys(value);
					} else if (action.equalsIgnoreCase("click")) {
						// scroll to view
						executor.executeScript("arguments[0].scrollIntoView();", element);
						executor.executeScript("arguments[0].click();", element);
					} else if (action.equalsIgnoreCase("isDisplayed")) {
						element.isDisplayed();
					} else if (action.equalsIgnoreCase("getText")) {
						String temp = "";
						String elementText = element.getText();
						temp = elementText;
						strQuote.append(temp + ",");
						System.out.println("Quote IDs : " + strQuote);

					} else if (action.equalsIgnoreCase("select")) {
						Select objSelect = new Select(element);
						objSelect.selectByVisibleText(value);

					}
					locatorType = null;
					break;

				case "name":
					element = driver.findElement(By.name(locatorValue));
					if (action.equalsIgnoreCase("sendkeys")) {
						element.clear();
						element.sendKeys(value);
					} else if (action.equalsIgnoreCase("click")) {
						// scroll to view
						executor.executeScript("arguments[0].scrollIntoView();", element);
						executor.executeScript("arguments[0].click();", element);
					} else if (action.equalsIgnoreCase("isDisplayed")) {
						element.isDisplayed();
					} else if (action.equalsIgnoreCase("getText")) {
						String temp = "";
						String elementText = element.getText();
						temp = elementText;
						strQuote.append(temp + ",");
						System.out.println("Quote IDs : " + strQuote);
					} else if (action.equalsIgnoreCase("select")) {
						Select objSelect = new Select(element);
						objSelect.selectByVisibleText(value);

					}
					locatorType = null;
					break;

				case "xpath":
					element = driver.findElement(By.xpath(locatorValue));
					if (action.equalsIgnoreCase("sendkeys")) {
						element.clear();
						element.sendKeys(value);
					} else if (action.equalsIgnoreCase("click")) {
						// scroll to view
						executor.executeScript("arguments[0].scrollIntoView();", element);
						executor.executeScript("arguments[0].click();", element);
					} else if (action.equalsIgnoreCase("isDisplayed")) {
						element.isDisplayed();
					} else if (action.equalsIgnoreCase("getText")) {
						String temp = "";
						String elementText = element.getText();
						temp = elementText;
						strQuote.append(temp + ",");
						System.out.println("Quote IDs : " + strQuote);

					} else if (action.equalsIgnoreCase("calendar date")) {
						StringBuffer elementText = new StringBuffer(element.toString());
						elementText.append("/div[text()='" + value + "']");
					} else if (action.equalsIgnoreCase("select")) {
						Select objSelect = new Select(element);
						objSelect.selectByVisibleText(value);

					}
					locatorType = null;
					break;

				case "cssSelector":
					element = driver.findElement(By.cssSelector(locatorValue));
					if (action.equalsIgnoreCase("sendkeys")) {
						element.clear();
						element.sendKeys(value);
					} else if (action.equalsIgnoreCase("click")) {
						// scroll to view
						executor.executeScript("arguments[0].scrollIntoView();", element);
						executor.executeScript("arguments[0].click();", element);
					} else if (action.equalsIgnoreCase("isDisplayed")) {
						element.isDisplayed();
					} else if (action.equalsIgnoreCase("getText")) {
						String temp = "";
						String elementText = element.getText();
						temp = elementText;
						strQuote.append(temp + ",");
						System.out.println("Quote IDs : " + strQuote);
					} else if (action.equalsIgnoreCase("select")) {
						Select objSelect = new Select(element);
						objSelect.selectByVisibleText(value);

					}
					locatorType = null;
					break;

				case "className":
					element = driver.findElement(By.className(locatorValue));
					if (action.equalsIgnoreCase("sendkeys")) {
						element.clear();
						element.sendKeys(value);
					} else if (action.equalsIgnoreCase("click")) {
						// scroll to view
						executor.executeScript("arguments[0].scrollIntoView();", element);
						executor.executeScript("arguments[0].click();", element);
					} else if (action.equalsIgnoreCase("isDisplayed")) {
						element.isDisplayed();
					} else if (action.equalsIgnoreCase("getText")) {
						String temp = "";
						String elementText = element.getText();
						temp = elementText;
						strQuote.append(temp + ",");
						System.out.println("Quote IDs : " + strQuote);
					} else if (action.equalsIgnoreCase("select")) {
						Select objSelect = new Select(element);
						objSelect.selectByVisibleText(value);

					}
					locatorType = null;
					break;

				case "linkText":
					element = driver.findElement(By.linkText(locatorValue));
					element.click();
					locatorType = null;
					break;

				case "partialLinkText":
					element = driver.findElement(By.partialLinkText(locatorValue));
					element.click();
					locatorType = null;
					break;

				default:
					break;
				}

			} catch (Exception e) {
				System.out.println(" Exception: " + e.toString());
			}

		}

	}
}
