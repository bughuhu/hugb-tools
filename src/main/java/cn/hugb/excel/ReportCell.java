package cn.hugb.excel;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;

public class ReportCell {
	public static final short ALIGN_GENERAL = 0;
	public static final short ALIGN_LEFT = 1;
	public static final short ALIGN_CENTER = 2;
	public static final short ALIGN_RIGHT = 3;
	public static final short ALIGN_FILL = 4;
	public static final short ALIGN_JUSTIFY = 5;
	public static final short ALIGN_CENTER_SELECTION = 6;
	public static final short VERTICAL_TOP = 0;
	public static final short VERTICAL_CENTER = 1;
	public static final short VERTICAL_BOTTOM = 2;
	public static final short VERTICAL_JUSTIFY = 3;
	public static final short BORDER_NONE = 0;
	public static final short BORDER_THIN = 1;
	public static final short BORDER_MEDIUM = 2;
	public static final short BORDER_DASHED = 3;
	public static final short BORDER_HAIR = 4;
	public static final short BORDER_THICK = 5;
	public static final short BORDER_DOUBLE = 6;
	public static final short BORDER_DOTTED = 7;
	public static final short BORDER_MEDIUM_DASHED = 8;
	public static final short BORDER_DASH_DOT = 9;
	public static final short BORDER_MEDIUM_DASH_DOT = 10;
	public static final short BORDER_DASH_DOT_DOT = 11;
	public static final short BORDER_MEDIUM_DASH_DOT_DOT = 12;
	public static final short BORDER_SLANTED_DASH_DOT = 13;
	public static final short NO_FILL = 0;
	public static final short SOLID_FOREGROUND = 1;
	public static final short FINE_DOTS = 2;
	public static final short ALT_BARS = 3;
	public static final short SPARSE_DOTS = 4;
	public static final short THICK_HORZ_BANDS = 5;
	public static final short THICK_VERT_BANDS = 6;
	public static final short THICK_BACKWARD_DIAG = 7;
	public static final short THICK_FORWARD_DIAG = 8;
	public static final short BIG_SPOTS = 9;
	public static final short BRICKS = 10;
	public static final short THIN_HORZ_BANDS = 11;
	public static final short THIN_VERT_BANDS = 12;
	public static final short THIN_BACKWARD_DIAG = 13;
	public static final short THIN_FORWARD_DIAG = 14;
	public static final short SQUARES = 15;
	public static final short DIAMONDS = 16;
	public static final short LESS_DOTS = 17;
	public static final short LEAST_DOTS = 18;
	public static final int FORMAT_CURRENCY = 0;
	public static final int FORMAT_CUSTOM = 3;
	public static final int FORMAT_DATE = 2;
	public static final int FORMAT_NUMBER = 1;
	public static final int FORMAT_STRING = 0;
	public static final int FORMAT_FORMULA = -1;
	private String cellContent;
	private HSSFWorkbook workbook;
	private HSSFSheet sheet;
	private ReportRow row;
	private int cellFormat;
	private int fontSize = 2000;
	private int hAlign = 0;
	private int vAlign = 1;
	private int borderWidth = 0;
	private boolean hasBorders;
	private boolean isBold = false;
	private boolean noBorder = false;
	private int columnWidth = 0;

	private String cellStyleName = "cellStyle";

	public HSSFCellStyle buildStyle(HSSFWorkbook wb) {
		HSSFCellStyle style = wb.createCellStyle();

		// 居中
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		if (isHasBorders()) {
			short borderWidth = (short) getBorderWidth();
			style.setBorderBottom(borderWidth);
			style.setBorderTop(borderWidth);
			style.setBorderLeft(borderWidth);
			style.setBorderRight(borderWidth);
			// 底部边框颜色
			style.setBottomBorderColor(new HSSFColor.BLACK().getIndex());
		}
		HSSFFont font = wb.createFont();
		font.setFontHeightInPoints((short) getFontSize());
		if (this.isBold) {
			font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		}
		style.setFont(font);
		return style;
	}

	public ReportCell(String cell, int cellFormat) {
		this.cellContent = cell;
		this.cellFormat = cellFormat;
	}

	public Integer getCellFormat() {
		return Integer.valueOf(this.cellFormat);
	}

	public void setCellFormat(int cellFormat) {
		this.cellFormat = cellFormat;
	}

	public String getCellContent() {
		return this.cellContent;
	}

	public void setCellContent(String cellContent) {
		this.cellContent = cellContent;
	}

	public boolean isBold() {
		return this.isBold;
	}

	public int getFontSize() {
		return this.fontSize;
	}

	public int getHAlign() {
		return this.hAlign;
	}

	public int getVAlign() {
		return this.vAlign;
	}

	public int getBorderWidth() {
		return this.borderWidth;
	}

	public void setBold(boolean isBold) {
		this.isBold = isBold;
	}

	public void setFontSize(int fontSize) {
		this.fontSize = fontSize;
	}

	public void setHAlign(int align) {
		this.hAlign = align;
	}

	public void setVAlign(int align) {
		this.vAlign = align;
	}

	public void setBorderWidth(int borderWidth) {
		this.borderWidth = borderWidth;
	}

	public boolean isHasBorders() {
		return this.hasBorders;
	}

	public void setHasBorders(boolean hasBorders) {
		this.hasBorders = hasBorders;
	}

	public HSSFWorkbook getWorkbook() {
		return this.workbook;
	}

	public HSSFSheet getSheet() {
		return this.sheet;
	}

	public ReportRow getRow() {
		return this.row;
	}

	public void setWorkbook(HSSFWorkbook workbook) {
		this.workbook = workbook;
	}

	public void setSheet(HSSFSheet sheet) {
		this.sheet = sheet;
	}

	public void setRow(ReportRow row) {
		this.row = row;
	}

	public String getCellStyleName() {
		return this.cellStyleName;
	}

	public void setCellStyleName(String cellStyleName) {
		this.cellStyleName = cellStyleName;
	}

	public boolean isNoBorder() {
		return this.noBorder;
	}

	public void setNoBorder(boolean noBorder) {
		this.noBorder = noBorder;
	}

	public int getColumnWidth() {
		return this.columnWidth;
	}

	public void setColumnWidth(int columnWidth) {
		this.columnWidth = columnWidth;
	}
}