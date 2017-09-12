package Functions;

import java.awt.AWTException;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;

import javax.imageio.ImageIO;

import org.apache.log4j.Logger;

public class ScreenShot {
	public void printScreen(String App, String Inst, int i, String SSFolder) {
		Logger log = Logger.getLogger("Logs");
		try {

			File file1 = new File(SSFolder);
			String name = App;
			String k = "" + i;
			String filename = name.concat("_").concat(k).concat("_").concat(Inst).concat(".png");
			BufferedImage image = new Robot()
					.createScreenCapture(new Rectangle(Toolkit.getDefaultToolkit().getScreenSize()));
			ImageIO.write(image, "png", new File(file1 + "\\" + filename));
			log.info("Screenshot Taken for Step#" + k + " with filename " + filename);

		} catch (AWTException e) {
			log.info(e);
		} catch (IOException e) {
			log.info(e);
		}
	}

}
