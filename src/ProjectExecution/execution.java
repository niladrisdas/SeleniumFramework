package ProjectExecution;
import java.awt.AWTException;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.image.BufferedImage;
import java.awt.event.KeyEvent;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;

import javax.imageio.ImageIO;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.apache.log4j.Logger;

public class execution {

	public void exeTC(String App, String Inst, String Xpath, String Url,
			WebDriver driver, String value) {
		Logger log = Logger.getLogger("Logs");
		if (Inst.equals("startApp")) {
			try {
				driver.get(Url);
				log.info("startApp Sucessfull");
			} catch (Exception e) {
				// System.out.println(e);
				log.info(e);

			}
		}
		else if (Inst.equals("type")) {

			driver.findElement(By.xpath(Xpath)).sendKeys(value);
			log.info("typing " + value + " Sucessfull");
		}
		else if (Inst.equals("click")) {
			driver.findElement(By.xpath(Xpath)).click();
			log.info("Click on " + Xpath + "Sucessfull");
		}
		else if (Inst.equals("Stop")) {
			// System.out.println("Execution completed");
			log.info("Execution completed");
			log.info("Teardown");
			System.exit(0);

		}
		else if (Inst.equals("wait")) {
			try {
				int time = Integer.parseInt(value);
				time = time * 1000;
				// System.out.println("Waitng for " + value + " secs");
				log.info("Waitng for " + value + " secs");
				System.out.println(" ");
				Thread.sleep(time);
			} catch (InterruptedException e) {
				// System.out.println(e);
				log.info(e);
			}
		}
		else if (Inst.equals("switchFrame")) {
			int index = Integer.parseInt(value);
			driver.switchTo().frame(index);
			// System.out.println(driver.getTitle());
			log.info(driver.getTitle());
		}
		else if (Inst.equals("doubleClick")) {
			Actions action = new Actions(driver);
			action.doubleClick(driver.findElement(By.xpath(Xpath))).build()
					.perform();
			log.info("doubleClick on " + Xpath + "Sucessfull");
		}
		else if (Inst.equals("handleWindows")) {
			try {
				java.util.Set<String> mainWindow = driver.getWindowHandles();
				// driver.switchTo().window(mainWindow);
				int windows = driver.getWindowHandles().size();
				// System.out.println(windows);
				log.info(windows);
				// System.out.println("Title Before handle " +
				// driver.getTitle());
				log.info("Title Before handle " + driver.getTitle());

				for (String otherWindow : driver.getWindowHandles()) {

					Thread.sleep(2000);
					for (int i = 1; i <= windows; i++) {
						if (!otherWindow.equals(mainWindow)) {
							driver.switchTo().window(otherWindow);
							/*
							 * System.out.println("Title after handle " +
							 * driver.getTitle());
							 */
							log.info("Title after handle " + driver.getTitle());
							System.out.println(" ");
							Thread.sleep(1000);
							break;
						}
					}
				}

			} catch (Exception e) {
				// System.out.println(e);
				log.info(e);
			}

		}
		else if (Inst.equals("verifyValue")) {
			String text = driver.findElement(By.xpath(Xpath)).getAttribute(
					"value");
			if (text.equals(value)) {
				// System.out.println("Text " + value + " Matches");
				log.info("Text " + value + " Matches");
				// System.out.println(" ");
			} else {
				// System.out.println("Expected:" + value + " Actual:" + text);
				log.info("Expected:" + value + " Actual:" + text);
				System.out.println(" ");
				System.exit(0);
			}
		}
		else if (Inst.equals("verifyText")) {
			String text = driver.findElement(By.xpath(Xpath)).getText();
			if (text.equals(value)) {
				// System.out.println("Text " + value + " Matches");
				log.info("Text " + value + " Matches");
				System.out.println(" ");
			} else {
				// System.out.println("Expected:" + value + " Actual:" + text);
				log.info("Expected:" + value + " Actual:" + text);
				System.out.println(" ");
				System.exit(0);
			}
		}
		else if (Inst.equals("switchCurrWindow")) {
			String parentWindow = driver.getWindowHandle();
			driver.switchTo().window(parentWindow);
			// System.out.println("Title after handle " + driver.getTitle());
			log.info("Title after handle " + driver.getTitle());
			System.out.println(" ");
		}
		else if (Inst.equals("selectFromDropdown")) {
			new Select(driver.findElement(By.xpath(Xpath)))
					.selectByVisibleText(value);
			log.info(value + "Selected");
		}
		else if (Inst.equals("pressKey")) {
			try {
				Robot key = new Robot();
				if (value.equals("DOWN")) {
					key.keyPress(KeyEvent.VK_DOWN);
					// System.out.println(value+" Key Pressed");
					log.info(value + " Key Pressed");
					Thread.sleep(1000);
					System.out.println(" ");
				} else if (value.equals("UP")) {
					key.keyPress(KeyEvent.VK_UP);
					// System.out.println(value+" Key Pressed");
					log.info(value + " Key Pressed");
					Thread.sleep(1000);
					System.out.println(" ");
				} else if (value.equals("LEFT")) {
					key.keyPress(KeyEvent.VK_LEFT);
					// System.out.println(value+" Key Pressed");
					log.info(value + " Key Pressed");
					Thread.sleep(1000);
					System.out.println(" ");
				} else if (value.equals("RIGHT")) {
					key.keyPress(KeyEvent.VK_RIGHT);
					// System.out.println(value+" Key Pressed");
					log.info(value + " Key Pressed");
					Thread.sleep(1000);
					System.out.println(" ");
				} else if (value.equals("ENTER")) {
					key.keyPress(KeyEvent.VK_ENTER);
					// System.out.println(value+" Key Pressed");
					log.info(value + " Key Pressed");
					Thread.sleep(1000);
					System.out.println(" ");
				}
			} catch (AWTException e) {
				// System.out.println(e);
				log.info(e);
			} catch (InterruptedException e) {
				// System.out.println(e);
				log.info(e);
			}
		}
		else if (Inst.equals("maximize")) {
					driver.manage().window().maximize();
		}
		else if(Inst.equals("waitUntil")){
			try{
			int time = Integer.parseInt(value);
			WebDriverWait wait=new WebDriverWait(driver,time);
			wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(Xpath)));
			log.info("Wait for Webelement"+Xpath+"Successfull");
			}catch(NoSuchElementException e){
				log.info(e);
			}
		}

		else if(Inst.equals("typeIfBlank")){
			WebElement element=driver.findElement(By.xpath(Xpath));
			if(element.getAttribute("value").isEmpty()){
				element.sendKeys(value);
			}else{
				log.info("field must be blank");
				//System.exit(0);

			}
		}

	}

	public void printScreen(String App, String Inst, int i,String SSFolder) {
		Logger log = Logger.getLogger("Logs");
		try {

			File file1 = new File(SSFolder);
			String name = App;
			String k = "" + i;
			String filename = name.concat("_").concat(k).concat(".png");
			BufferedImage image = new Robot()
					.createScreenCapture(new Rectangle(Toolkit
							.getDefaultToolkit().getScreenSize()));
			ImageIO.write(image, "png", new File(file1 + "\\" + filename));

		} catch (AWTException e) {
			log.info(e);
		} catch (IOException e) {
			log.info(e);
		}
	}

	@SuppressWarnings("deprecation")
	public void readXcel(String filepath, String filename,String SSFolder,String IEDriver) throws IOException {
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
			String app = null, inst = null, xpath = null, url = null, ss = null;
			String value = null;
			WebDriver driver = null;
			DataFormatter formatter = new DataFormatter();
			Sheet sh = wb.getSheetAt(1);
			// String url=null,app,inst,xpath;
			int j = 1;
			int k = 1;

			int row = sh.getLastRowNum() - sh.getFirstRowNum();
			for (int i = 1; i < row + 1; i++) {

				app = sh.getRow(i).getCell(0,org.apache.poi.ss.usermodel.Row.CREATE_NULL_AS_BLANK ).getStringCellValue();
				inst = sh.getRow(i).getCell(2,org.apache.poi.ss.usermodel.Row.CREATE_NULL_AS_BLANK ).getStringCellValue();
				xpath = sh.getRow(i).getCell(3,org.apache.poi.ss.usermodel.Row.CREATE_NULL_AS_BLANK ).getStringCellValue();
				ss = sh.getRow(i).getCell(7,org.apache.poi.ss.usermodel.Row.CREATE_NULL_AS_BLANK ).getStringCellValue();
				Cell cell = sh.getRow(i).getCell(4,org.apache.poi.ss.usermodel.Row.CREATE_NULL_AS_BLANK );
				value = formatter.formatCellValue(cell);
				// System.out.println("App=" + app);
				log.info("App=" + app);
				// System.out.println("Instruction=" + inst);
				log.info("Instruction=" + inst);
				// System.out.println("xpath=" + xpath);
				log.info("xpath=" + xpath);

				// System.out.println("SS=" + ss);
				System.out.println(" ");

				if (inst.equals("startApp")) {
					Sheet sh1 = wb.getSheetAt(0);
					 int row1 = sh1.getLastRowNum() - sh1.getFirstRowNum();
				for(j=1;j<=row1;j++){
					if ((sh.getRow(i).getCell(0).getStringCellValue())
							.equals(sh1.getRow(j).getCell(0)
									.getStringCellValue())) {
						url = sh1.getRow(j).getCell(1).getStringCellValue();
						// System.out.print("URL=" + url);
						log.info("URL=" + url);
						System.out.println(" ");
						Runtime.getRuntime().exec(
								"taskkill /F /IM IEDriverServer.exe");
						File file1 = new File(IEDriver);
						System.setProperty("webdriver.ie.driver",
								file1.getAbsolutePath());
						driver = new InternetExplorerDriver();
						//j++;
						break;
					}
				}log.info(url);
				} else if (url.equals("null")){
					//url = null;
					log.info("No URL found");
					System.exit(0);

				}
                if(inst!=""){
				exeTC(app, inst, xpath, url, driver, value);}
                else{
                	log.info("No More Steps to Execute");
                	System.exit(0);
                }


               if (ss.equals("Y")) {
					printScreen(app, inst, k,SSFolder);
					k++;
				}

			}
		} catch (IOException e) {
			// System.out.println(e);
			log.error(e);
		}
	}

	public static void main(String args[]) throws FileNotFoundException {
		String path = "C:/Users/niladri.das/Downloads/NTP/Selenium/Framework/config/config.txt";
		final Logger log = Logger.getLogger("Logs");
		FileReader fr = new FileReader(path);
		BufferedReader br = new BufferedReader(fr);
		String data;
		String path1 = null;
		String name = null;
		String SSFolder=null;
		String IEDriver=null;
		try {

			while ((data = br.readLine()) != null) {
				if (data.substring(0, 8).equals("filepath")) {
					path1 = data.substring(9);
					// System.out.println("Filepath:" + path1);
					log.info("Filepath:" + path1);
				} else if (data.substring(0, 8).equals("filename")) {
					name = data.substring(9);
					// System.out.println("Filename:" + name);
					log.info("Filename:" + name);
					System.out.println(" ");
				}else if(data.substring(0,16).equals("ScreenShotFolder")){
					SSFolder=data.substring(17);
					log.info("SreenshotFolDerPath:" + SSFolder);
				}else if(data.substring(0,12).equals("IEDriverPath")){
					IEDriver=data.substring(13);
					log.info("IEDriverPath:" + IEDriver);
				}else{
					break;
				}
			}
			execution obj = new execution();
			obj.readXcel(path1, name,SSFolder,IEDriver);

		} catch (Exception e) {
			// System.out.println(e);
			log.error(e);
		}
	}
}
