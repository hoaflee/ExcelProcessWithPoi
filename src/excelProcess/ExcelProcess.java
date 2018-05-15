package excelProcess;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Iterator;
import java.util.List;
import java.util.logging.FileHandler;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.logging.SimpleFormatter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelProcess {
	private final static Logger logger = Logger.getLogger(ExcelProcess.class.getName());
	private static FileHandler fh = null;

	public static String GetExecutionPath() {
		String path = "";
		try {
			String executionPath = System.getProperty("user.dir");
			path = executionPath.replace("\\", "/");
		} catch (Exception e) {
			System.out.println("Exception caught =" + e.getMessage());
		}
		return path;
	}

	public static void main(String[] args) {
		String src = "";
		String des = "";

		for (int i = 0; i < args.length - 1; i++) {
			if (args[i].equalsIgnoreCase("-src")) {
				src = args[i + 1];
			}
			if (args[i].equalsIgnoreCase("-des")) {
				des = args[i + 1];
			}
		}
		if (src.isEmpty()) {
			System.out.println("arguments must be supplied");
			System.out.println("Usage: java -jar excelcopy.jar -src <arg0> -des <arg1>");
			System.exit(1);
		}
		if (des.isEmpty()) {
			System.out.println("arguments must be supplied");
			System.out.println("Usage: java -jar excelcopy.jar -src <arg0> -des <arg1>");
			System.exit(1);
		}

		SimpleDateFormat format = new SimpleDateFormat("M-d_HHmmss");
		try {
			fh = new FileHandler(GetExecutionPath() + File.separator + "ExcelCopy-"
					+ format.format(Calendar.getInstance().getTime()) + ".log");
		} catch (Exception e) {
			e.printStackTrace();
		}

		fh.setFormatter(new SimpleFormatter());
		logger.addHandler(fh);

		List<String> excelFiles = parseDir(src);
		for (String file : excelFiles) {
			try {
				logger.info("******* " + file + " *********");
				FileInputStream excelFile = new FileInputStream(new File(file));
				Workbook workbook = new XSSFWorkbook(excelFile);

				// 生年月日不整合 sheet
				insertExcelSheetName(des, workbook, 0);
				// 死亡記載あり sheet
				insertExcelSheetName(des, workbook, 1);
				// メモあり sheet
				insertExcelSheetName(des, workbook, 2);

				// ４、情報不整合 sheet
				insertExcelContent(des, workbook, 3);
				// ５、情報不明瞭 sheet
				insertExcelContent(des, workbook, 4);

			} catch (FileNotFoundException e) {
				logger.log(Level.SEVERE, e.getMessage());
			} catch (IOException e) {
				logger.log(Level.SEVERE, e.getMessage());
			}
		}
	}

	/**
	 * Get all excel filr from source dir
	 * @param dirPath
	 * @return
	 */
	public static List<String> parseDir(String dirPath) {
		List<String> excelFiles = new ArrayList<String>();
		File folder = new File(dirPath);
		if (folder.isDirectory()) {
			File files[] = folder.listFiles();
			for (File file : files) {
				if (file.getName().endsWith(".xlsx")) {
					excelFiles.add(file.getAbsolutePath());
				}
			}
		}
		return excelFiles;
	}

	private static void insertExcelContent(String desDir, Workbook workbook, int sheetIndex) {
		String sheetName = workbook.getSheetName(sheetIndex);
		logger.info("=============== START UPDATE AT SHEET " + sheetName + " ===============");
		Sheet sheet = workbook.getSheetAt(sheetIndex);
		Iterator<Row> rows = sheet.iterator();

		while (rows.hasNext()) {
			Row currenRow = rows.next();
			if (currenRow.getRowNum() == 0) {
				continue; // skip column header
			}
			Cell currentCell = currenRow.getCell(0);
			String cellValue = currentCell.getStringCellValue().trim();
			String replaceContent = currenRow.getCell(1).getStringCellValue();
			if (!cellValue.isEmpty()) {
				insertExcel(desDir, cellValue, replaceContent);
			}
		}
	}

	private static void insertExcelSheetName(String desDir, Workbook workbook, int sheetIndex) {
		String sheetName = workbook.getSheetName(sheetIndex);
		logger.info("=============== START UPDATE AT SHEET " + sheetName + " ===============");

		Sheet sheet = workbook.getSheetAt(sheetIndex);
		Iterator<Row> rows = sheet.iterator();

		while (rows.hasNext()) {
			Row currenRow = rows.next();
			if (currenRow.getRowNum() == 0) {
				continue; // skip column header
			}
			Cell currentCell = currenRow.getCell(0);
			String cellValue = currentCell.getStringCellValue().trim();
			if (!cellValue.isEmpty()) {
				insertExcel(desDir, cellValue, sheetName);
			}
		}
	}

	/**
	 * Update excel file with insertString
	 *
	 * @param excelFileName
	 * @param insertString
	 */
	private static void insertExcel(String desDir, String excelFileName, String insertString) {

		String newContent = "";
		Cell cell = null;

		// search file in forder
		FileSearch fileSearch = new FileSearch();

		// try different directory and filename :)
		fileSearch.searchDirectory(new File(desDir), excelFileName + ".xlsx");

		int count = fileSearch.getResult().size();
		if (count == 0) {
			logger.log(Level.SEVERE, excelFileName + ".xlsx No file found!");
		} else {
			for (String matched : fileSearch.getResult()) {
				try {
					FileInputStream excelFile = new FileInputStream(new File(matched));
					Workbook workbook = new XSSFWorkbook(excelFile);
					Sheet sheet = workbook.getSheetAt(0);
					cell = sheet.getRow(1).getCell(5);
					String currentValue = cell.getStringCellValue();

					if (!currentValue.isEmpty()) {
						newContent = currentValue + "\n" + insertString;
						CellStyle style = cell.getCellStyle(); // Create new style
						style.setWrapText(true); // Set wordwrap
						cell.setCellStyle(style);
					} else {
						newContent = insertString;
					}
					cell.setCellValue(newContent);
					excelFile.close();

					FileOutputStream outFile = new FileOutputStream(new File(matched));
					workbook.write(outFile);
					outFile.close();

					logger.info(matched + " has been updated.");
				} catch (FileNotFoundException e) {
					logger.log(Level.SEVERE, e.getMessage());
				} catch (IOException e) {
					logger.log(Level.SEVERE, e.getMessage());
				}
			}
		}

	}
}
