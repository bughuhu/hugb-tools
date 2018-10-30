package cn.hugb.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {
	private String filePath;
	private HSSFWorkbook workbook;
	private XSSFWorkbook workbookx;

	public ExcelReader() {
	}

	public ExcelReader(String file) throws FileNotFoundException, IOException {
		this.filePath = file;

		if (file.endsWith(".xls"))
			this.workbook = new HSSFWorkbook(new FileInputStream(this.filePath));
		if (file.endsWith(".xlsx"))
			this.workbookx = new XSSFWorkbook(new FileInputStream(this.filePath));
	}

	public void loadFile(String file) throws FileNotFoundException, IOException {
		this.filePath = file;

		if (file.endsWith(".xls"))
			this.workbook = new HSSFWorkbook(new FileInputStream(this.filePath));
		if (file.endsWith(".xlsx"))
			this.workbookx = new XSSFWorkbook(new FileInputStream(this.filePath));
	}

	public String getCellValue(int sheet, int row, int column) {
		if (this.filePath.endsWith(".xlsx")) {
			return getCellValueX(sheet, row, column);
		}
		HSSFSheet childSheet = this.workbook.getSheetAt(sheet);
		if (childSheet == null)
			return "";
		HSSFRow objRow = childSheet.getRow(row);
		if (objRow == null)
			return "";
		HSSFCell cell = childSheet.getRow(row).getCell(column);
		if (cell == null) {
			return "";
		}
		String value = "";

		switch (cell.getCellType()) {
		case 0:
			if (HSSFDateUtil.isCellDateFormatted(cell)) {
				SimpleDateFormat dateformat = new SimpleDateFormat("yyyy-MM-dd");
				Date dt = HSSFDateUtil.getJavaDate(cell.getNumericCellValue());
				value = dateformat.format(dt);
			} else {
				double d = cell.getNumericCellValue();
				DecimalFormat f = new DecimalFormat("0.000000");
				value = subZeroAndDot(f.format(d));
			}
			break;
		case 1:
			value = cell.getStringCellValue();
			break;
		case 2:
		case 3:
		}

		return value;
	}

	private String getCellValueX(int sheet, int row, int column) {
		XSSFSheet childSheet = this.workbookx.getSheetAt(sheet);
		if (childSheet == null)
			return "";
		XSSFRow objRow = childSheet.getRow(row);
		if (objRow == null)
			return "";
		XSSFCell cell = objRow.getCell(column);

		String value = "";

		if (cell == null)
			return "";
		switch (cell.getCellType()) {
		case 0:
			if (HSSFDateUtil.isCellDateFormatted(cell)) {
				SimpleDateFormat dateformat = new SimpleDateFormat("yyyy-MM-dd");
				Date dt = HSSFDateUtil.getJavaDate(cell.getNumericCellValue());
				value = dateformat.format(dt);
			} else {
				double d = cell.getNumericCellValue();
				DecimalFormat f = new DecimalFormat("0.000000");
				value = subZeroAndDot(f.format(d));
			}
			break;
		case 1:
			value = cell.getStringCellValue();
			break;
		case 3:
			value = "";
			break;
		case 2:
		default:
			value = "";
		}

		return value;
	}

	public String subZeroAndDot(String s) {
		if (s.indexOf(".") > 0) {
			s = s.replaceAll("0+?$", "");
			s = s.replaceAll("[.]$", "");
		}
		return s;
	}

	public static void main(String[] argv) throws FileNotFoundException, IOException {
		String fileToBeRead = "D:/test.xlsx";
		ExcelReader excelReader = new ExcelReader();
		excelReader.loadFile(fileToBeRead);
		String cell = excelReader.getCellValue(0, 1, 8);
		System.out.println(cell);
		cell = excelReader.getCellValue(0, 2, 8);
		System.out.println(cell);
	}
}