package Functions;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;

public class ReadExcel {
	WriteExcel we = new WriteExcel();
	public void readSteps(String userAction, String filepath, String filename, String SSFolder, String IEDriver)
			throws Exception {
		Logger log = Logger.getLogger("Logs");
		try {

			Workbook wb = null;
			File file = new File(filepath + "\\" + filename);
			FileInputStream inputStream = new FileInputStream(file);
			String fileExtn = filename.substring(filename.indexOf("."));
			if (fileExtn.equals(".xlsx")) {
				wb = new XSSFWorkbook(inputStream);
			} else if (fileExtn.equals(".xls")) {
				wb = new HSSFWorkbook(inputStream);
			}
			wb.setMissingCellPolicy(MissingCellPolicy.CREATE_NULL_AS_BLANK);
			String app = null, inst = null, xpath = null, url = null, ss = null;
			String value = null;
			String ua;
			WebDriver driver = null;
			DataFormatter formatter = new DataFormatter();
			Sheet sh = wb.getSheetAt(1);
			int j;
			int k = 1;
			int row = sh.getLastRowNum() - sh.getFirstRowNum();

			for (int i = 1; i <= row; i++) {
				ua = sh.getRow(i).getCell(1).getStringCellValue();
				while (ua.equals(userAction)) {
					app = sh.getRow(i).getCell(0)
							.getStringCellValue();
					inst = sh.getRow(i).getCell(3)
							.getStringCellValue();
					xpath = sh.getRow(i).getCell(4)
							.getStringCellValue();
					ss = sh.getRow(i).getCell(8)
							.getStringCellValue();
					Cell cell = sh.getRow(i).getCell(5);
					value = formatter.formatCellValue(cell);
					log.info("App=" + app);
					log.info("Instruction=" + inst);
					log.info("xpath=" + xpath);
					System.out.println(" ");

					if (inst.equals("startApp")) {
						Sheet sh1 = wb.getSheetAt(0);
						int row1 = sh1.getLastRowNum() - sh1.getFirstRowNum();
						for (j = 1; j <= row1; j++) {
							if ((sh.getRow(i).getCell(0).getStringCellValue())
									.equals(sh1.getRow(j).getCell(0).getStringCellValue())) {
								url = sh1.getRow(j).getCell(1).getStringCellValue();
								log.info("URL=" + url);
								System.out.println(" ");
								// Runtime.getRuntime().exec("taskkill /F /IM IEDriverServer.exe");
								File file1 = new File(IEDriver);
								System.setProperty("webdriver.ie.driver", file1.getAbsolutePath());
								driver = new InternetExplorerDriver();
								break;
							}
						}
						log.info(url);
					} else if (url.equals("null")) {
						log.info("No URL found");
						System.exit(0);

					}
					if (inst != "") {
						BaseFunctions bs = new BaseFunctions();
						// WriteExcel we = new WriteExcel();
						// we.setStartTime(filepath, filename, i);
						bs.exeTC(app, inst, xpath, url, driver, value, i, filepath, filename);
						// we.setEndTime(filepath, filename, i);
						// we.calculateTotalTime(filepath, filename, i);
					} else {
						log.info("No More Steps to Exceute");
						System.exit(0);
					}
					if (ss.equals("Y")) {
						ScreenShot scr = new ScreenShot();
						scr.printScreen(app, inst, k, SSFolder);
						k++;
					} else {
						k++;
					}
					break;
				}

			}

		} catch (IOException e) {
			log.error(e);
		}
	}

	public void readUserAction(String filepath, String filename, String ssFolder, String ieDriverPath) {
		/* Build In Progress */
		Logger log = Logger.getLogger("Logs");
		try {
			Workbook wb1 = null;
			File file = new File(filepath + "\\" + filename);
			FileInputStream inputStream = new FileInputStream(file);
			String fileExtn = filename.substring(filename.indexOf("."));
			if (fileExtn.equals(".xlsx")) {
				wb1 = new XSSFWorkbook(inputStream);
			} else if (fileExtn.equals(".xls")) {
				wb1 = new HSSFWorkbook(inputStream);
			}
			wb1.setMissingCellPolicy(MissingCellPolicy.CREATE_NULL_AS_BLANK);
			String ua = "";
			String path;
			String name;
			Sheet sh = wb1.getSheetAt(0);
			int row;
			String results;
			int totRow = sh.getLastRowNum() - sh.getFirstRowNum();
			for (row = 1; row < totRow + 1; row++) {
				ua = sh.getRow(row).getCell(0).getStringCellValue();
				path = sh.getRow(row).getCell(1).getStringCellValue();
				name = sh.getRow(row).getCell(2).getStringCellValue();

				if (ua != "") {
					log.info("Executing User Action " + ua);
					try {
						we.setStartTimeTc(filepath, filename, row);
						readSteps(ua, path, name, ssFolder, ieDriverPath);
						results = "Passed";
						we.writeXcelTc(filepath, filename, results, row);
						we.setEndTimeTc(filepath, filename, row);
						we.calculateTotalTimeTc(filepath, filename, row);
					} catch (Exception e) {
						log.info(e);
						results = "Failed";
						we.writeXcelTc(filepath, filename, results, row);
					}
				}

			}

		} catch (Exception e) {
			log.info(e);
		}

	}
}
