package Functions;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
//import java.io.IOException;

import org.apache.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

public class BaseFunctions {
	public void exeTC(String App, String Inst, String Xpath, String Url, WebDriver driver, String value, int rowNum,
			String filename, String filePath) throws Exception {
		Logger log = Logger.getLogger("Logs");
		WriteExcel we = new WriteExcel();
		String result;
		if (Inst.equals("startApp")) {
			try {
				driver.get(Url);
				log.info("startApp Sucessfull");
				result = "Passed";
				we.writeXcel(filename, filePath, result, rowNum);
			} catch (Exception e) {
				log.info(e);
				result = "Failed";
				we.writeXcel(filename, filePath, result, rowNum);
				System.exit(0);
			}
		} else if (Inst.equals("type")) {
			try {
				driver.findElement(By.xpath(Xpath)).sendKeys(value);
				log.info("typing " + value + " Sucessfull");
				result = "Passed";
				we.writeXcel(filename, filePath, result, rowNum);
			} catch (Exception e) {
				log.info(e);
				result = "Failed";
				we.writeXcel(filename, filePath, result, rowNum);
				System.exit(0);
			}
		} else if (Inst.equals("click")) {
			try {
				driver.findElement(By.xpath(Xpath)).click();
				log.info("Click on " + Xpath + "Sucessfull");
				result = "Passed";
				we.writeXcel(filename, filePath, result, rowNum);
			} catch (Exception e) {
				log.info(e);
				result = "Failed";
				we.writeXcel(filename, filePath, result, rowNum);
				System.exit(0);
			}
		} else if (Inst.equals("Stop")) {
			log.info("Execution completed");
			log.info("Teardown");
			System.exit(0);

		} else if (Inst.equals("wait")) {
			try {
				int time = Integer.parseInt(value);
				time = time * 1000;
				log.info("Waitng for " + value + " secs");
				System.out.println(" ");
				Thread.sleep(time);
				result = "Passed";
				we.writeXcel(filename, filePath, result, rowNum);
			} catch (InterruptedException e) {
				log.info(e);
				result = "Failed";
				we.writeXcel(filename, filePath, result, rowNum);
				System.exit(0);
			}
		} else if (Inst.equals("switchFrame")) {
			try {
				int index = Integer.parseInt(value);
				driver.switchTo().frame(index);
				log.info(driver.getTitle());
				result = "Passed";
				we.writeXcel(filename, filePath, result, rowNum);
			} catch (Exception e) {
				log.info(e);
				result = "Failed";
				we.writeXcel(filename, filePath, result, rowNum);
				System.exit(0);
			}
		} else if (Inst.equals("doubleClick")) {
			try {
				Actions action = new Actions(driver);
				action.doubleClick(driver.findElement(By.xpath(Xpath))).build().perform();
				log.info("doubleClick on " + Xpath + "Sucessfull");
				result = "Passed";
				we.writeXcel(filename, filePath, result, rowNum);
			} catch (Exception e) {
				log.info(e);
				result = "Failed";
				we.writeXcel(filename, filePath, result, rowNum);
				System.exit(0);
			}
		} else if (Inst.equals("handleWindows")) {
			try {
				java.util.Set<String> mainWindow = driver.getWindowHandles();

				int windows = driver.getWindowHandles().size();
				log.info(windows);
				log.info("Title Before handle " + driver.getTitle());

				for (String otherWindow : driver.getWindowHandles()) {

					Thread.sleep(2000);
					for (int i = 1; i <= windows; i++) {
						if (!otherWindow.equals(mainWindow)) {
							driver.switchTo().window(otherWindow);
							log.info("Title after handle " + driver.getTitle());
							System.out.println(" ");
							Thread.sleep(1000);
							break;
						}
					}
				}
				result = "Passed";
				we.writeXcel(filename, filePath, result, rowNum);
			} catch (Exception e) {
				log.info(e);
				result = "Failed";
				we.writeXcel(filename, filePath, result, rowNum);
				System.exit(0);

			}

		} else if (Inst.equals("verifyValue")) {
			String text = driver.findElement(By.xpath(Xpath)).getAttribute("value");
			if (text.equals(value)) {
				log.info("Text " + value + " Matches");
				result = "Passed";
				we.writeXcel(filename, filePath, result, rowNum);
			} else {
				log.info("Expected:" + value + " Actual:" + text);
				System.out.println(" ");
				result = "Failed";
				we.writeXcel(filename, filePath, result, rowNum);
				System.exit(0);
			}
		} else if (Inst.equals("verifyText")) {
			String text = driver.findElement(By.xpath(Xpath)).getText();
			if (text.equals(value)) {
				log.info("Text " + value + " Matches");
				result = "Passed";
				we.writeXcel(filename, filePath, result, rowNum);
			} else {
				log.info("Expected:" + value + " Actual:" + text);
				result = "Failed";
				we.writeXcel(filename, filePath, result, rowNum);
				System.exit(0);
			}
		} else if (Inst.equals("switchCurrWindow")) {
			try {
				String parentWindow = driver.getWindowHandle();
				driver.switchTo().window(parentWindow);
				log.info("Title after handle " + driver.getTitle());
				result = "Passed";
				we.writeXcel(filename, filePath, result, rowNum);
			} catch (Exception e) {
				log.info(e);
				result = "Failed";
				we.writeXcel(filename, filePath, result, rowNum);
				System.exit(0);
			}
		} else if (Inst.equals("selectFromDropdown")) {
			try {
				new Select(driver.findElement(By.xpath(Xpath))).selectByVisibleText(value);
				log.info(value + "Selected");
				result = "Passed";
				we.writeXcel(filename, filePath, result, rowNum);
			} catch (Exception e) {
				log.info(e);
				result = "Failed";
				we.writeXcel(filename, filePath, result, rowNum);
				System.exit(0);
			}
		} else if (Inst.equals("pressKey")) {
			try {
				Robot key = new Robot();
				if (value.equals("DOWN")) {
					key.keyPress(KeyEvent.VK_DOWN);
					log.info(value + " Key Pressed");
					Thread.sleep(1000);
					System.out.println(" ");
				} else if (value.equals("UP")) {
					key.keyPress(KeyEvent.VK_UP);
					log.info(value + " Key Pressed");
					Thread.sleep(1000);
					System.out.println(" ");
				} else if (value.equals("LEFT")) {
					key.keyPress(KeyEvent.VK_LEFT);
					log.info(value + " Key Pressed");
					Thread.sleep(1000);
					System.out.println(" ");
				} else if (value.equals("RIGHT")) {
					key.keyPress(KeyEvent.VK_RIGHT);
					log.info(value + " Key Pressed");
					Thread.sleep(1000);
					System.out.println(" ");
				} else if (value.equals("ENTER")) {
					key.keyPress(KeyEvent.VK_ENTER);
					log.info(value + " Key Pressed");
					Thread.sleep(1000);
					System.out.println(" ");
				}
				result = "Passed";
				we.writeXcel(filename, filePath, result, rowNum);
			} catch (AWTException e) {
				log.info(e);
				result = "Failed";
				we.writeXcel(filename, filePath, result, rowNum);
				System.exit(0);
			} catch (InterruptedException e) {
				log.info(e);
				result = "Failed";
				we.writeXcel(filename, filePath, result, rowNum);
				System.exit(0);
			}
		} else if (Inst.equals("maximize")) {
			try {
				driver.manage().window().maximize();
				result = "Passed";
				we.writeXcel(filename, filePath, result, rowNum);
			} catch (Exception e) {
				log.info(e);
				result = "Failed";
				we.writeXcel(filename, filePath, result, rowNum);
				System.exit(0);
			}
		} else if (Inst.equals("waitUntil")) {
			try {
				int time = Integer.parseInt(value);
				WebDriverWait wait = new WebDriverWait(driver, time);
				wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(Xpath)));
				log.info("Wait for Webelement" + Xpath + "Successfull");
				result = "Passed";
				we.writeXcel(filename, filePath, result, rowNum);
			} catch (NoSuchElementException e) {
				log.info(e);
				result = "Failed";
				we.writeXcel(filename, filePath, result, rowNum);
				System.exit(0);
			}
		}

		else if (Inst.equals("typeIfBlank")) {
			WebElement element = driver.findElement(By.xpath(Xpath));
			if (element.getAttribute("value").isEmpty()) {
				element.sendKeys(value);
				result = "Passed";
				we.writeXcel(filename, filePath, result, rowNum);
			} else {
				log.info("field must be blank");
				result = "Failed";
				we.writeXcel(filename, filePath, result, rowNum);
				System.exit(0);

			}
		}

	}
}
