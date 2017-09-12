package ProjectExecution;


import java.io.File;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.sikuli.script.App;
import org.sikuli.script.FindFailed;
import org.sikuli.script.Pattern;
import org.sikuli.script.Screen;

public class sikuli {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		try{
		Runtime.getRuntime().exec(
				"taskkill /F /IM IEDriverServer.exe");
		File file1 = new File("D:\\Exec-Environment\\iAF6_1-Test\\webdrivers\\IEDriverServer.exe");
		System.setProperty("webdriver.ie.driver",
				file1.getAbsolutePath());
		WebDriver driver = new InternetExplorerDriver();
		driver.get("www.google.com");
		
		Screen screen = new Screen();
		Pattern img=new Pattern("D:\\Encrypt\\Selenuim Framework\\png\\textbox.png");
		screen.click(img);
		}catch(Exception e){
			
		}
	}

}
