package cn.hugb.excel;

import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFFooter;
import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.hssf.usermodel.HSSFPrintSetup;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

public class ReportDefinition {
	private String sheetName = "sheet1";
	private String reportName = "";
	private String operatorName = "";

	private List<ReportParameter> parametes = new ArrayList();
	private List<ReportRow> headers = new ArrayList();
	private List<ReportRow> rows = new ArrayList();
	private HSSFWorkbook workbook;
	private HSSFSheet sheet;
	private boolean prepared = false;
	private HSSFCellStyle headerStyle;
	private HSSFCellStyle propertyStyle;
	private HSSFCellStyle cellStyle;
	private HSSFCellStyle stringCellStyle;
	private HSSFCellStyle numberCellStyle;
	private HSSFCellStyle dateCellStyle;
	private HSSFCellStyle currencyStyle;
	private HSSFCellStyle percentageStyle;
	public static String HEADER_STYLE = "headerStyle";
	public static String CELL_STYLE = "cellStyle";
	public static String STRING_STYLE = "stringStyle";
	public static String NUMBER_STYLE = "numberStyle";
	public static String DATE_STYLE = "dateStyle";
	public static String PROPERTY_STYLE = "propertyStyle";
	public static String CURRENCY_STYLE = "currencyStyle";
	public static String PERCENTAGE_STYLE = "percentageStyle";

	private HashMap<String, HSSFCellStyle> styles = new HashMap();

	private boolean printProperty = true;
	private int headerRowIndex = 0;

	private int dataStartRowIndex = 0;
	private int dataEndRowIndex = 0;
	private int dataStartColIndex = 0;
	private int dataEndColIndex = 0;

	private int titleStartRowIndex = 0;
	private int titleEndRowIndex = 0;
	private int titleStartColIndex = 0;
	private int titleEndColIndex = 0;
	private boolean printable = false;

	private SimpleDateFormat fm = new SimpleDateFormat("yyyy-MM-dd HH:mm:dd");

	public HSSFCellStyle getBoldStyle(HSSFCell cell) {
		HSSFCellStyle cellStyle = this.workbook.createCellStyle();
		cellStyle.cloneStyleFrom(cell.getCellStyle());
		HSSFFont font = this.workbook.createFont();
		font.setBoldweight((short) 700);
		cellStyle.setFont(font);
		return cellStyle;
	}

	public void setCellBold(int rowIndex, int colIndex) {
		if (this.sheet.getRow(rowIndex) != null) {
			HSSFRow row = this.sheet.getRow(rowIndex);
			if (row.getCell(colIndex) != null) {
				HSSFCell cell = this.sheet.getRow(rowIndex).getCell(colIndex);
				cell.setCellStyle(getBoldStyle(cell));
			}
		}
	}

	public void setBottomLine(int rowIndex, int colIndex, short linewidth) {
		if (this.sheet.getRow(rowIndex) != null) {
			HSSFRow row = this.sheet.getRow(rowIndex);
			if (row.getCell(colIndex) != null) {
				HSSFCell cell = this.sheet.getRow(rowIndex).getCell(colIndex);

				HSSFCellStyle cellStyle = this.workbook.createCellStyle();
				cellStyle.cloneStyleFrom(cell.getCellStyle());
				cellStyle.setBorderBottom(linewidth);

				cell.setCellStyle(getBoldStyle(cell));
			}
		}
	}

	public HSSFCellStyle getNoBorderStyle(HSSFCell cell) {
		HSSFCellStyle cellStyle = this.workbook.createCellStyle();
		cellStyle.cloneStyleFrom(cell.getCellStyle());
		cellStyle.setBorderLeft((short) 0);
		cellStyle.setBorderRight((short) 0);
		cellStyle.setBorderTop((short) 0);
		cellStyle.setBorderBottom((short) 0);
		return cellStyle;
	}

	public void setBorder(HSSFCell cell, short width) {
		HSSFCellStyle cellStyle = cell.getCellStyle();
		cellStyle.setBorderLeft(width);
		cellStyle.setBorderRight(width);
		cellStyle.setBorderTop(width);
		cellStyle.setBorderBottom(width);
	}

	public void setBorder(int rowIndex, int colIndex, short width) {
		if (this.sheet.getRow(rowIndex) != null) {
			HSSFRow row = this.sheet.getRow(rowIndex);
			if (row.getCell(colIndex) != null) {
				HSSFCell cell = this.sheet.getRow(rowIndex).getCell(colIndex);
				setBorder(cell, width);
			}
		}
	}

	public ReportDefinition() {
		this.workbook = new HSSFWorkbook();
		this.sheet = this.workbook.createSheet(this.sheetName);
	}

	private void prepareStyles() {
		this.headerStyle = this.workbook.createCellStyle();
		if (!this.printable) {
			this.headerStyle.setBorderLeft((short) 1);
			this.headerStyle.setBorderRight((short) 1);
			this.headerStyle.setBorderTop((short) 1);
			this.headerStyle.setBorderBottom((short) 1);
		}
		HSSFFont headerFont = this.workbook.createFont();
		headerFont.setBoldweight((short) 700);
		this.headerStyle.setFont(headerFont);
		this.headerStyle.setVerticalAlignment((short) 1);
		this.headerStyle.setAlignment((short) 2);
		this.headerStyle.setWrapText(true);

		this.cellStyle = this.workbook.createCellStyle();
		if (!this.printable) {
			this.cellStyle.setBorderLeft((short) 1);
			this.cellStyle.setBorderRight((short) 1);
			this.cellStyle.setBorderTop((short) 1);
			this.cellStyle.setBorderBottom((short) 1);
		}
		this.cellStyle.setVerticalAlignment((short) 1);
		this.cellStyle.setAlignment((short) 2);

		this.propertyStyle = this.workbook.createCellStyle();
		if (!this.printable) {
			this.propertyStyle.setBorderLeft((short) 1);
			this.propertyStyle.setBorderRight((short) 1);
			this.propertyStyle.setBorderTop((short) 1);
			this.propertyStyle.setBorderBottom((short) 1);
		}
		this.propertyStyle.setFont(headerFont);
		this.propertyStyle.setVerticalAlignment((short) 1);
		this.propertyStyle.setAlignment((short) 1);

		this.stringCellStyle = this.workbook.createCellStyle();
		if (!this.printable) {
			this.stringCellStyle.setBorderLeft((short) 1);
			this.stringCellStyle.setBorderRight((short) 1);
			this.stringCellStyle.setBorderTop((short) 1);
			this.stringCellStyle.setBorderBottom((short) 1);
		}
		this.stringCellStyle.setVerticalAlignment((short) 1);
		this.stringCellStyle.setAlignment((short) 1);

		this.numberCellStyle = this.workbook.createCellStyle();
		if (!this.printable) {
			this.numberCellStyle.setBorderLeft((short) 1);
			this.numberCellStyle.setBorderRight((short) 1);
			this.numberCellStyle.setBorderTop((short) 1);
			this.numberCellStyle.setBorderBottom((short) 1);
		}
		this.numberCellStyle.setVerticalAlignment((short) 1);
		this.numberCellStyle.setAlignment((short) 3);

		HSSFDataFormat format = this.workbook.createDataFormat();
		this.dateCellStyle = this.workbook.createCellStyle();
		this.dateCellStyle.setDataFormat(format.getFormat("yyyy-MM-dd"));
		if (!this.printable) {
			this.dateCellStyle.setBorderLeft((short) 1);
			this.dateCellStyle.setBorderRight((short) 1);
			this.dateCellStyle.setBorderTop((short) 1);
			this.dateCellStyle.setBorderBottom((short) 1);
		}
		this.dateCellStyle.setVerticalAlignment((short) 1);
		this.dateCellStyle.setAlignment((short) 3);

		this.currencyStyle = this.workbook.createCellStyle();
		HSSFDataFormat currencyformat = this.workbook.createDataFormat();
		this.currencyStyle.setDataFormat(currencyformat.getFormat("#,##0.00"));
		if (!this.printable) {
			this.currencyStyle.setBorderLeft((short) 1);
			this.currencyStyle.setBorderRight((short) 1);
			this.currencyStyle.setBorderTop((short) 1);
			this.currencyStyle.setBorderBottom((short) 1);
		}
		this.currencyStyle.setVerticalAlignment((short) 1);
		this.currencyStyle.setAlignment((short) 3);

		this.percentageStyle = this.workbook.createCellStyle();
		HSSFDataFormat percentageFormat = this.workbook.createDataFormat();
		this.percentageStyle.setDataFormat(percentageFormat.getFormat("0.00%"));
		if (!this.printable) {
			this.percentageStyle.setBorderLeft((short) 1);
			this.percentageStyle.setBorderRight((short) 1);
			this.percentageStyle.setBorderTop((short) 1);
			this.percentageStyle.setBorderBottom((short) 1);
		}
		this.percentageStyle.setVerticalAlignment((short) 1);
		this.percentageStyle.setAlignment((short) 3);

		this.styles.put(HEADER_STYLE, this.headerStyle);
		this.styles.put(PROPERTY_STYLE, this.propertyStyle);
		this.styles.put(STRING_STYLE, this.stringCellStyle);
		this.styles.put(CELL_STYLE, this.cellStyle);
		this.styles.put(NUMBER_STYLE, this.numberCellStyle);
		this.styles.put(DATE_STYLE, this.dateCellStyle);
		this.styles.put(CURRENCY_STYLE, this.currencyStyle);
		this.styles.put(PERCENTAGE_STYLE, this.percentageStyle);
	}

	public void addRow(ReportRow row) {
		row.setRowNumber(5 + this.parametes.size() + 1 + this.headers.size() + this.rows.size());
		this.rows.add(row);
	}

	public void addHeader(String header) {
		if (this.headers.size() == 0) {
			this.headers.add(new ReportRow());
		}
		((ReportRow) this.headers.get(0)).addCell(header);
	}

	public void addHeaderWidth(String header, int width) {
		if (this.headers.size() == 0) {
			this.headers.add(new ReportRow());
		}
		((ReportRow) this.headers.get(0)).addCellWidth(header, width);
	}

	public void addHeaderRow(ReportRow row) {
		this.headers.add(row);
	}

	public void addParameter(String propertyName, String propertyValue) {
		this.parametes.add(new ReportParameter(propertyName, propertyValue));
	}

	public void mergeCells(int firstRow, int lastRow, int firstCol, int lastCol) {
		CellRangeAddress address = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
		this.workbook.getSheet(this.sheetName).addMergedRegion(address);
	}

	public String getCellValue(int rowIndex, int colIndex) {
		String value = "";
		HSSFSheet sheet = this.workbook.getSheet(this.sheetName);
		if (sheet.getRow(rowIndex) == null) {
			return null;
		}
		HSSFRow row = sheet.getRow(rowIndex);
		if (row.getCell(colIndex) == null) {
			return null;
		}
		HSSFCell cell = row.getCell(colIndex);
		switch (cell.getCellType()) {
		case 0:
			value = Double.toString(cell.getNumericCellValue());
			break;
		case 1:
			value = cell.getStringCellValue();
			break;
		case 2:
		case 3:
		}

		return value;
	}

	public void mergeSameCellsInRow(int theRow, int firstCol, int lastCol) {
		if (!this.prepared) {
			prepareExcel();
		}
		int fromCol = firstCol;
		int toCol = firstCol + 1;
		for (int i = firstCol + 1; i <= lastCol; i++) {
			toCol = i;
			String v1 = getCellValue(theRow, fromCol);
			String v2 = getCellValue(theRow, toCol);
			if ((v1 == null) || (v2 == null)) {
				toCol--;
			} else if (!v1.equals(v2)) {
				if (fromCol + 1 != toCol)
					mergeCells(theRow, theRow, fromCol, toCol - 1);
				fromCol = toCol;
			}
		}
		mergeCells(theRow, theRow, fromCol, toCol);
	}

	public void mergeSameCellsInColumn(int theCol, int firstRow, int lastRow) {
		if (!this.prepared) {
			prepareExcel();
		}
		int fromRow = firstRow;
		int toRow = firstRow + 1;
		for (int i = firstRow + 1; i <= lastRow; i++) {
			toRow = i;
			String v1 = getCellValue(fromRow, theCol);
			String v2 = getCellValue(toRow, theCol);
			if ((v1 == null) || (v2 == null)) {
				toRow--;
			} else if (!v1.equals(v2)) {
				if (fromRow + 1 != toRow)
					mergeCells(fromRow, toRow - 1, theCol, theCol);
				fromRow = toRow;
			}
		}
		mergeCells(fromRow, toRow, theCol, theCol);
	}

	public int getFirstHeaderRow() {
		return this.parametes.size();
	}

	public int getLastHeaderRow() {
		return this.parametes.size() + this.headers.size() - 1;
	}

	public int getFirstContentRow() {
		return this.parametes.size() + this.headers.size();
	}

	public int getLastContentRow() {
		return this.parametes.size() + this.headers.size() + this.rows.size() - 1;
	}

	public int getLastColumn() {
		if (this.headers.size() == 0)
			return 0;
		return ((ReportRow) this.headers.get(0)).size() - 1;
	}

	private void insertPropertyCells(int rownumber, String paramName, String paraValue, boolean merge) {
		HSSFRow row = this.sheet.createRow(rownumber);
		HSSFCell cellPropertyName = row.createCell(0);
		cellPropertyName.setCellStyle(this.propertyStyle);
		cellPropertyName.setCellType(1);
		cellPropertyName.setCellValue(paramName);

		HSSFCell cellPropertyValue = row.createCell(1);
		cellPropertyValue.setCellStyle(this.stringCellStyle);
		cellPropertyValue.setCellType(1);
		cellPropertyValue.setCellValue(paraValue);

		if (merge)
			mergeCells(rownumber, rownumber, 0, 1);
	}

	public void prepareExcel() {
		if (!this.prepared) {
			prepareStyles();
			int rowNumber = 0;
			if (this.printProperty) {
				insertPropertyCells(rowNumber++, "报表信息", "", true);
				insertPropertyCells(rowNumber++, "报表名称", getReportName(), false);
				insertPropertyCells(rowNumber++, "制表人", getOperatorName(), false);
				insertPropertyCells(rowNumber++, "制表时间", this.fm.format(new Date()), false);
				insertPropertyCells(rowNumber++, "报表参数", "", true);

				for (ReportParameter parameter : this.parametes) {
					insertPropertyCells(rowNumber++, parameter.getPropertyName(), parameter.getPropertyValue(), false);
				}
				rowNumber++;
			}

			this.dataStartRowIndex = rowNumber;
			this.titleStartRowIndex = rowNumber;

			for (int h = 0; h < this.headers.size(); h++) {
				HSSFRow row = this.sheet.createRow(rowNumber++);
				for (int i = 0; i < ((ReportRow) this.headers.get(h)).size(); i++) {
					HSSFCell cellHeader = ((HSSFRow) row).createCell(i);
					cellHeader.setCellStyle(this.headerStyle);
					cellHeader.setCellType(1);
					cellHeader.setCellValue(
							((ReportCell) ((ReportRow) this.headers.get(h)).getCells().get(i)).getCellContent());
					this.titleEndColIndex = i;
				}
			}
			this.titleEndRowIndex = (rowNumber - 1);

			for (Object row = this.rows.iterator(); ((Iterator) row).hasNext();) {
				ReportRow reportRow = (ReportRow) ((Iterator) row).next();
				HSSFRow row1 = this.sheet.createRow(rowNumber++);
				for (int i = 0; i < reportRow.getCells().size(); i++) {
					HSSFCell cell = row1.createCell(i);

					ReportCell rc = (ReportCell) reportRow.getCells().get(i);
					String cellValue = rc.getCellContent();
					Integer cellFormat = rc.getCellFormat();

					cell.setCellType(1);
					cell.setCellStyle((HSSFCellStyle) this.styles.get(rc.getCellStyleName()));

					if (cellFormat.intValue() == -1) {
						cell.setCellFormula(rc.getCellContent());
						cell.setCellType(2);
					}

					if ((cellFormat.intValue() == 2) && (isDate(cellValue))) {
						try {
							SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
							cell.setCellValue(sdf.parse(cellValue));
						} catch (Exception ex) {
							cell.setCellType(1);
							cell.setCellValue(cellValue);
						}
					}
					if (cellFormat.intValue() == 1) {
						try {
							if (isInteger(cellValue)) {
								if (cellValue.length() < 11)
									cell.setCellValue(Integer.parseInt(cellValue));
								cell.setCellType(0);
							} else if (isDecimal(cellValue)) {
								cell.setCellValue(Double.parseDouble(cellValue));
								cell.setCellType(0);
							}
						} catch (Exception ex) {
							cell.setCellValue(cellValue);
							cell.setCellType(1);
						}
					}
					if (cellFormat.intValue() == 0) {
						cell.setCellValue(cellValue);
						cell.setCellType(1);
					}
					if (cellFormat.intValue() == 3) {
						cell.setCellType(1);
						HSSFCellStyle customStyle = rc.buildStyle(this.workbook);
						cell.setCellValue(rc.getCellContent());
						cell.setCellStyle(customStyle);
					}
					if (rc.isBold()) {
						cell.setCellStyle(getBoldStyle(cell));
					}

					if (rc.isNoBorder()) {
						cell.setCellStyle(getNoBorderStyle(cell));
					}
				}
				row1.setHeightInPoints(row1.getHeightInPoints() * 1.2F);
			}

			this.dataEndRowIndex = (rowNumber - 1);
			this.dataStartColIndex = 0;

			if (this.headers.size() > 0) {
				for (int i = 0; i < ((ReportRow) this.headers.get(0)).size(); i++) {
					this.sheet.autoSizeColumn(i);

					ReportCell rc = (ReportCell) ((ReportRow) this.headers.get(0)).getCells().get(i);
					if (rc.getColumnWidth() == 0)
						this.sheet.setColumnWidth(i, (int) (this.sheet.getColumnWidth(i) * 1.3D));
					else {
						this.sheet.setColumnWidth(i, rc.getColumnWidth());
					}
					this.dataEndColIndex = i;
				}
			} else if (this.rows.size() > 0) {
				for (int i = 0; i < ((ReportRow) this.rows.get(0)).size(); i++) {
					this.sheet.autoSizeColumn(i);

					ReportCell rc = (ReportCell) ((ReportRow) this.rows.get(0)).getCells().get(i);
					if (rc.getColumnWidth() == 0)
						this.sheet.setColumnWidth(i, (int) (this.sheet.getColumnWidth(i) * 1.3D));
					else {
						this.sheet.setColumnWidth(i, rc.getColumnWidth());
					}
					this.dataEndColIndex = i;
				}
			}

			if (this.headers.size() == 1) {
				this.sheet.getRow(this.titleStartRowIndex)
						.setHeightInPoints(this.sheet.getRow(this.titleStartRowIndex).getHeightInPoints() * 2.0F);
			}

			if (this.printable) {
				addSignFooter(rowNumber);
			}
		}
		this.prepared = true;
	}

	public void buildExcel(String filePath) {
		try {
			if (!this.prepared) {
				prepareExcel();
			}

			FileOutputStream fOut = new FileOutputStream(filePath);
			this.workbook.write(fOut);
			fOut.flush();
			fOut.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void setFooter(int fontSize, String footerStr) {
		HSSFFooter footer = getSheet().getFooter();
		footer.setCenter(HSSFFooter.fontSize((short) fontSize) + footerStr);
	}

	public void setFooterCenter(String footerCenter) {
		HSSFFooter footer = getSheet().getFooter();
		footer.setCenter(footerCenter);
	}

	public void setPageNumberSizeAndFooter(int fontSize, String str) {
		HSSFFooter footer = getSheet().getFooter();
		str = str.replaceAll("#PageNumber#", HSSFFooter.page());
		str = str.replaceAll("#PageCount#", HSSFFooter.numPages());
		footer.setRight(HSSFFooter.fontSize((short) fontSize) + str);
	}

	public void setPageNumberFooter() {
		HSSFFooter footer = getSheet().getFooter();
		footer.setRight("第" + HSSFFooter.page() + "页 共" + HSSFFooter.numPages() + "页");
	}

	public void addSignFooter(int rowNumber) {
	}

	public boolean isDecimal(String str) {
		if ((str == null) || ("".equals(str))) {
			return false;
		}
		Pattern pattern = Pattern.compile("^(-?\\d+)(\\.\\d+)?");

		return pattern.matcher(str).matches();
	}

	public boolean isInteger(String str) {
		if (str == null)
			return false;
		Pattern pattern = Pattern.compile("[0-9]+");
		return pattern.matcher(str).matches();
	}

	public boolean isDate(String str) {
		if (str == null)
			return false;
		Pattern pattern = Pattern
				.compile("^([1-2]\\d{3})[\\/|\\-](0?[1-9]|10|11|12)[\\/|\\-]([1-2]?[0-9]|0[1-9]|30|31)$");
		return pattern.matcher(str).matches();
	}

	public List<ReportParameter> getParametes() {
		return this.parametes;
	}

	public List<ReportRow> getHeaders() {
		return this.headers;
	}

	public List<ReportRow> getRows() {
		return this.rows;
	}

	public void setParametes(List<ReportParameter> parametes) {
		this.parametes = parametes;
	}

	public void setHeaders(List<ReportRow> headers) {
		this.headers = headers;
	}

	public void setRows(List<ReportRow> rows) {
		this.rows = rows;
	}

	public String getSheetName() {
		return this.sheetName;
	}

	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	public String getReportName() {
		return this.reportName;
	}

	public String getOperatorName() {
		return this.operatorName;
	}

	public void setReportName(String reportName) {
		this.reportName = reportName;
	}

	public void setOperatorName(String operatorName) {
		this.operatorName = operatorName;
	}

	public boolean isPrintProperty() {
		return this.printProperty;
	}

	public void setPrintProperty(boolean printProperty) {
		this.printProperty = printProperty;
	}

	public int getHeaderRowIndex() {
		return this.headerRowIndex;
	}

	public void setHeaderRowIndex(int headerRowIndex) {
		this.headerRowIndex = headerRowIndex;
	}

	public void setColumnWidth(int columnIndex, int width) {
		this.sheet.setColumnWidth(columnIndex, width);
	}

	public void setRowHeight(int rowIndex, int height) {
		this.sheet.getRow(rowIndex).setHeight((short) height);
	}

	public HSSFWorkbook getWorkbook() {
		return this.workbook;
	}

	public HSSFSheet getSheet() {
		return this.sheet;
	}

	public void setWorkbook(HSSFWorkbook workbook) {
		this.workbook = workbook;
	}

	public void setSheet(HSSFSheet sheet) {
		this.sheet = sheet;
	}

	public void hideRow(int rowIndex) {
		if (!this.prepared)
			prepareExcel();
		this.sheet.getRow(rowIndex).setZeroHeight(true);
	}

	public void setDefaultHeader() {
		HSSFHeader header = getSheet().getHeader();
		header.setCenter(getReportName());
	}

	public void setHeader(int fontSize, String headerStr) {
		HSSFHeader header = getSheet().getHeader();
		header.setCenter(HSSFHeader.fontSize((short) fontSize) + headerStr);
	}

	public void setDefaultPrintArea() {
		getWorkbook().setPrintArea(0, this.dataStartColIndex, this.dataEndColIndex, this.dataStartRowIndex,
				this.dataEndRowIndex);
	}

	public void setDefaultRepeatTitle() {
		getWorkbook().setRepeatingRowsAndColumns(0, this.titleStartColIndex, this.titleEndColIndex,
				this.titleStartRowIndex, this.titleEndRowIndex);
	}

	public void setPageMargin(double left, double top, double right, double bottom) {
		getSheet().setMargin((short) 0, left);
		getSheet().setMargin((short) 1, right);
		getSheet().setMargin((short) 2, top);
		getSheet().setMargin((short) 3, bottom);
	}

	public void setDefaultFooter() {
		HSSFFooter footer = getSheet().getFooter();
		String footLeft = "";
		footLeft = footLeft + "制表人签字: __________________   ";
		footLeft = footLeft + "CB负责人签字: __________________   ";
		footLeft = footLeft + "HR 负责人签字: __________________   ";
		footLeft = footLeft + "子公司负责人签字: __________________";
		footer.setCenter(footLeft);

		footer.setRight("第" + HSSFFooter.page() + "页 共" + HSSFFooter.numPages() + "页");
	}

	public void setLandScape(boolean ls) {
		getSheet().getPrintSetup().setLandscape(ls);
	}

	public void setPageSize(int pagesize) {
		getSheet().getPrintSetup().setPaperSize((short) pagesize);
	}

	public HSSFPrintSetup getPrintSetup() {
		return getSheet().getPrintSetup();
	}

	public void hideColumn(int colIndex) {
		getSheet().setColumnHidden(colIndex, true);
	}

	public void setOrder(int colIndex) {
		for (int i = 0; i < this.rows.size(); i++)
			((ReportRow) this.rows.get(i)).setOrder(colIndex);
	}

	public void sort() {
		ReportRowComparator comparator = new ReportRowComparator();
		Collections.sort(this.rows, comparator);
	}

	public int getDataStartRowIndex() {
		return this.dataStartRowIndex;
	}

	public int getDataEndRowIndex() {
		return this.dataEndRowIndex;
	}

	public int getDataStartColIndex() {
		return this.dataStartColIndex;
	}

	public int getDataEndColIndex() {
		return this.dataEndColIndex;
	}

	public int getTitleStartRowIndex() {
		return this.titleStartRowIndex;
	}

	public int getTitleEndRowIndex() {
		return this.titleEndRowIndex;
	}

	public int getTitleStartColIndex() {
		return this.titleStartColIndex;
	}

	public int getTitleEndColIndex() {
		return this.titleEndColIndex;
	}

	public boolean isPrintable() {
		return this.printable;
	}

	public void setDataStartRowIndex(int dataStartRowIndex) {
		this.dataStartRowIndex = dataStartRowIndex;
	}

	public void setDataEndRowIndex(int dataEndRowIndex) {
		this.dataEndRowIndex = dataEndRowIndex;
	}

	public void setDataStartColIndex(int dataStartColIndex) {
		this.dataStartColIndex = dataStartColIndex;
	}

	public void setDataEndColIndex(int dataEndColIndex) {
		this.dataEndColIndex = dataEndColIndex;
	}

	public void setTitleStartRowIndex(int titleStartRowIndex) {
		this.titleStartRowIndex = titleStartRowIndex;
	}

	public void setTitleEndRowIndex(int titleEndRowIndex) {
		this.titleEndRowIndex = titleEndRowIndex;
	}

	public void setTitleStartColIndex(int titleStartColIndex) {
		this.titleStartColIndex = titleStartColIndex;
	}

	public void setTitleEndColIndex(int titleEndColIndex) {
		this.titleEndColIndex = titleEndColIndex;
	}

	public void setPrintable(boolean printable) {
		this.printable = printable;
	}

	public static void main(String[] args) {
		ReportDefinition definition = new ReportDefinition();

		definition.setOperatorName("LIUKAIHUA");
		definition.setReportName("示例表");

		ReportRow headerRow2 = new ReportRow();
		headerRow2.addCell("子公司");
		headerRow2.addCell("一级部门");
		headerRow2.addCell("二级部门");
		headerRow2.addCell("三级部门");
		headerRow2.addCell("四级部门");
		headerRow2.addCell("正式员工");
		headerRow2.addCellWidth("正式员工调整", 12000);
		headerRow2.addCell("正式员工小计");
		headerRow2.addCell("派遣\n员工");
		headerRow2.addCell("派遣员工\n调整");
		headerRow2.addCell("派遣员工\n小计");
		headerRow2.addCell("总计");
		headerRow2.addCell("正式\n员工");
		headerRow2.addCell("正式员工\n调整");
		headerRow2.addCell("正式员工\n小计");
		headerRow2.addCell("派遣员工");
		headerRow2.addCell("派遣员工\n调整");
		headerRow2.addCell("派遣员工小计");
		headerRow2.addCell("总计");

		definition.addHeaderRow(headerRow2);

		ReportRow dataRow1 = new ReportRow();
		dataRow1.addString("集团总部");
		dataRow1.addString("搜狐技术部15");
		dataRow1.addString("");
		dataRow1.addString("");
		dataRow1.addDate("2012-03-01");
		dataRow1.addNumber("150");
		dataRow1.addNumber("150");
		dataRow1.addNumber("150");
		dataRow1.addNumber("150");
		dataRow1.addNumber("150");
		dataRow1.addNumber("150");
		dataRow1.addNumber("150");
		dataRow1.addCurrencyNumber("15001230.23");
		dataRow1.addCurrencyNumber("0");
		dataRow1.addCurrencyNumber("150");
		dataRow1.addPercentNumber("150");
		dataRow1.addPercentNumber("1.50");
		dataRow1.addPercentNumber("25.0");
		dataRow1.addNumber("150");
		definition.addRow(dataRow1);
		for (int i = 1; i < 100; i++) {
			ReportRow dataRow2 = new ReportRow();
			dataRow2.addBoldString("集团总部");
			dataRow2.addCell("搜狐技术部" + i);
			dataRow2.addCell("");
			dataRow2.addCell("");
			dataRow2.addCell("");
			dataRow2.addCell("150");
			dataRow2.addCell("150");
			dataRow2.addCell("150");
			dataRow2.addCurrencyNumber("0.00");
			dataRow2.addCurrencyNumber("150150150.15");
			dataRow2.addCurrencyNumber("150.13");
			dataRow2.addCell("150");
			dataRow2.addCell("150");
			dataRow2.addCell("150");
			dataRow2.addCell("150");
			dataRow2.addNoborderText("150");
			dataRow2.addCell("150");
			dataRow2.addFormulaString("SUM(F10:Q10)");
			dataRow2.addCustomCell("Hello", true, 3, 2, 150, false, 28);

			definition.addRow(dataRow2);
		}

		ReportRow dataRow3 = new ReportRow();
		dataRow3.addString("集团总部");
		dataRow3.addString("搜狐技术部15");
		dataRow3.addString("");
		dataRow3.addString("");
		dataRow3.addDate("2012-03-01");
		dataRow3.sumAbove(7, 106, true);
		dataRow3.sumAbove(7, 106, true);
		dataRow3.sumAbove(7, 106, true);
		dataRow3.sumAbove(7, 106, true);
		dataRow3.sumAbove(7, 106, true);
		dataRow3.sumAbove(7, 106, true);
		dataRow3.addCurrencyNumber("15001230.23");
		dataRow3.addCurrencyNumber("0");
		dataRow3.addCurrencyNumber("150");
		dataRow3.addPercentNumber("150");
		dataRow3.addPercentNumber("1.50");
		dataRow3.addPercentNumber("25.0");
		dataRow3.addNumber("150");
		definition.addRow(dataRow3);

		definition.prepareExcel();
		definition.setPageNumberSizeAndFooter(16, "第#PageNumber#页 共#PageCount#页");
		definition.getWorkbook().setRepeatingRowsAndColumns(0, 0, 18, 6, 7);
		definition.buildExcel("D:\\sample.xls");
	}
}