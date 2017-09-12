package ProjectExecution;

import Functions.*;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import org.apache.log4j.Logger;

public class StartFramework {

	private static BufferedReader br;

	public static void main(String args[]) throws FileNotFoundException {
		File file = new File("C:/Users/niladri.das/Downloads/NTP/Selenium/Framework/Config/config.txt");
		final Logger log = Logger.getLogger("Logs");
		FileReader fr = new FileReader(file);
		br = new BufferedReader(fr);
		String data;
		String path1 = null;
		String name = null;
		String SSFolder = null;
		String Driver = null;
		try {

			while ((data = br.readLine()) != null) {
				if (data.substring(0, 8).equals("filepath")) {
					path1 = data.substring(9);
					log.info("Filepath:" + path1);
				} else if (data.substring(0, 8).equals("filename")) {
					name = data.substring(9);
					log.info("Filename:" + name);
					System.out.println(" ");
				} else if (data.substring(0, 11).equals("BrowserType")) {
					String Type = data.substring(12);
					log.info("Browser Type:" + Type);
					while ((data = br.readLine()) != null) {
						if (Type.equals("InternetExplorer")) {
							if (data.substring(0, 12).equals("IEDriverPath")) {
								Driver = data.substring(13);
								log.info("IEDriverPath:" + SSFolder);
								break;
							}
						} else if (Type.equals("Chrome")) {
							if (data.substring(0, 16).equals("ChromeDriverPath")) {
								Driver = data.substring(17);
								log.info("ChromeDriverPath:" + Driver);
								break;
							}
						} else {
							break;
						}
					}
				} else if (data.substring(0, 16).equals("ScreenShotFolder")) {
					SSFolder = data.substring(17);
					log.info("SreenshotFolDerPath:" + SSFolder);
				} else {
					break;
				}
			}
			ReadExcel re = new ReadExcel();
			re.readUserAction(path1, name, SSFolder, Driver);

		} catch (Exception e) {
			log.error(e);
		}
	}
}
