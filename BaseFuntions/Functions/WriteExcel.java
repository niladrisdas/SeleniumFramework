package Functions;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jfree.util.Log;

public class WriteExcel {

	public long time1;
	public long time2;
	public long time3;
	public long time4;

	public void writeXcel(String filepath, String filename, String result, int rowNum) throws IOException {
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
			Sheet sh = wb.getSheetAt(1);
			Row r = sh.getRow(rowNum);
			Cell c = r.getCell(9);
			c.setCellValue(result);
			if (result == "Passed") {
				XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
				style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
				style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				c.setCellStyle(style);
			} else if (result == "Failed") {
				XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
				style.setFillForegroundColor(IndexedColors.RED.getIndex());
				style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				c.setCellStyle(style);
			}

			inputStream.close();
			FileOutputStream outputStream = new FileOutputStream(file);
			wb.write(outputStream);

		} catch (IOException e) {
			log.error(e);
		}
	}

	public void writeXcelTc(String filepath, String filename, String result, int rowNum) throws IOException {
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
			Sheet sh = wb.getSheetAt(0);
			Row r = sh.getRow(rowNum);
			Cell c = r.getCell(4);
			c.setCellValue(result);
			if (result == "Passed") {
				XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
				style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
				style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				c.setCellStyle(style);
			} else if (result == "Failed") {
				XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
				style.setFillForegroundColor(IndexedColors.RED.getIndex());
				style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				c.setCellStyle(style);
			}

			inputStream.close();
			FileOutputStream outputStream = new FileOutputStream(file);
			wb.write(outputStream);

		} catch (IOException e) {
			log.error(e);
		}
	}

	public void setStartTime(String filepath, String filename, int rowNum) {
		try {
			time1 = System.currentTimeMillis();
			Date date = new Date();
			SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy h:mm:ss a");
			String formattedDate = sdf.format(date);
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
			Sheet sh = wb.getSheetAt(1);
			Row r = sh.getRow(rowNum);
			Cell c = r.getCell(10);
			c.setCellValue(formattedDate);
			inputStream.close();
			FileOutputStream outputStream = new FileOutputStream(file);
			wb.write(outputStream);
		} catch (Exception e) {
			Log.info(e);
		}

	}

	public void setStartTimeTc(String filepath, String filename, int rowNum) {
		try {
			time3 = System.currentTimeMillis();
			Date date = new Date();
			SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy h:mm:ss a");
			String formattedDate = sdf.format(date);
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
			Sheet sh = wb.getSheetAt(0);
			Row r = sh.getRow(rowNum);
			Cell c = r.getCell(5);
			c.setCellValue(formattedDate);
			inputStream.close();
			FileOutputStream outputStream = new FileOutputStream(file);
			wb.write(outputStream);
		} catch (Exception e) {
			Log.info(e);
		}

	}

	public void setEndTime(String filepath, String filename, int rowNum) {

		try {
			time2 = System.currentTimeMillis();
			Date date = new Date();
			SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy h:mm:ss a");
			String formattedDate = sdf.format(date);
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
			Sheet sh = wb.getSheetAt(1);
			Row r = sh.getRow(rowNum);
			Cell c = r.getCell(11);
			c.setCellValue(formattedDate);
			inputStream.close();
			FileOutputStream outputStream = new FileOutputStream(file);
			wb.write(outputStream);
		} catch (Exception e) {
			Log.info(e);
		}

	}

	public void setEndTimeTc(String filepath, String filename, int rowNum) {

		try {
			time4 = System.currentTimeMillis();
			Date date = new Date();
			SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy h:mm:ss a");
			String formattedDate = sdf.format(date);
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
			Sheet sh = wb.getSheetAt(0);
			Row r = sh.getRow(rowNum);
			Cell c = r.getCell(6);
			c.setCellValue(formattedDate);
			inputStream.close();
			FileOutputStream outputStream = new FileOutputStream(file);
			wb.write(outputStream);
		} catch (Exception e) {
			Log.info(e);
		}

	}

	public void calculateTotalTime(String filepath, String filename, int rowNum) {
		try {
			long totalTime = time2 - time1;
			totalTime = totalTime / 1000;
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

			Sheet sh = wb.getSheetAt(1);
			Row r = sh.getRow(rowNum);
			Cell c = r.getCell(12);
			System.out.println(c);
			c.setCellValue(totalTime + " secs");
			inputStream.close();
			FileOutputStream outputStream = new FileOutputStream(file);
			wb.write(outputStream);
		} catch (Exception e) {
			Log.info(e);
		}

	}

	public void calculateTotalTimeTc(String filepath, String filename, int rowNum) {
		try {
			long totalTime = time4 - time3;
			totalTime = totalTime / 1000;
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

			Sheet sh = wb.getSheetAt(0);
			Row r = sh.getRow(rowNum);
			Cell c = r.getCell(7);
			System.out.println(c);
			c.setCellValue(totalTime + " secs");
			inputStream.close();
			FileOutputStream outputStream = new FileOutputStream(file);
			wb.write(outputStream);
		} catch (Exception e) {
			Log.info(e);
		}

	}

}
