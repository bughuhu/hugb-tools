package cn.hugb.excel;

import java.util.ArrayList;
import java.util.List;

/**
 * 
 * @author huguobiao
 *
 */
public class ReportRow
{
  private List<ReportCell> cells = new ArrayList<ReportCell>();
  private int rowNumber;
  private String sortKey;

  public ReportRow()
  {
    this.sortKey = "";
  }

  public void addHeaderCell(String cell)
  {
    ReportCell rc = new ReportCell(cell, 0);
    rc.setCellStyleName(ReportDefinition.HEADER_STYLE);
    rc.setRow(this);
    this.cells.add(rc);
  }

  public void addString(String cell)
  {
    ReportCell rc = new ReportCell(cell, 0);
    rc.setCellStyleName(ReportDefinition.STRING_STYLE);
    rc.setRow(this);
    this.cells.add(rc);
  }

  public void addNumber(String cell)
  {
    ReportCell rc = new ReportCell(cell, 1);
    rc.setCellStyleName(ReportDefinition.NUMBER_STYLE);
    rc.setRow(this);
    this.cells.add(rc);
  }

  public void addDate(String cell)
  {
    ReportCell rc = new ReportCell(cell, 2);
    rc.setCellStyleName(ReportDefinition.DATE_STYLE);
    rc.setRow(this);
    this.cells.add(rc);
  }

  public void addCell(String cell)
  {
    addString(cell);
  }

  public void addCellWidth(String cell, int width)
  {
    ReportCell rc = new ReportCell(cell, 0);
    rc.setCellStyleName(ReportDefinition.STRING_STYLE);
    rc.setRow(this);
    rc.setColumnWidth(width);
    this.cells.add(rc);
  }

  public void addBoldString(String cell)
  {
    ReportCell rc = new ReportCell(cell, 0);
    rc.setCellStyleName(ReportDefinition.STRING_STYLE);
    rc.setBold(true);
    rc.setRow(this);
    this.cells.add(rc);
  }

  public void addCurrencyNumber(String cell)
  {
    ReportCell rc = new ReportCell(cell, 1);
    rc.setCellStyleName(ReportDefinition.CURRENCY_STYLE);
    rc.setRow(this);
    this.cells.add(rc);
  }

  public void addBoldCurrencyNumber(String cell)
  {
    ReportCell rc = new ReportCell(cell, 1);
    rc.setCellStyleName(ReportDefinition.CURRENCY_STYLE);
    rc.setBold(true);
    rc.setRow(this);
    this.cells.add(rc);
  }

  public void addPercentNumber(String cell)
  {
    ReportCell rc = new ReportCell(cell, 1);
    rc.setCellStyleName(ReportDefinition.PERCENTAGE_STYLE);
    rc.setRow(this);
    this.cells.add(rc);
  }

  public void addBoldDate(String cell)
  {
    ReportCell rc = new ReportCell(cell, 2);
    rc.setCellStyleName(ReportDefinition.DATE_STYLE);
    rc.setBold(true);
    rc.setRow(this);
    this.cells.add(rc);
  }

  public void addBoldNumber(String cell)
  {
    ReportCell rc = new ReportCell(cell, 1);
    rc.setCellStyleName(ReportDefinition.NUMBER_STYLE);
    rc.setBold(true);
    rc.setRow(this);
    this.cells.add(rc);
  }

  public void addNoborderBoldText(String cell)
  {
    ReportCell rc = new ReportCell(cell, 0);
    rc.setCellStyleName(ReportDefinition.STRING_STYLE);
    rc.setBold(true);
    rc.setNoBorder(true);
    rc.setRow(this);
    this.cells.add(rc);
  }

  public void addNoborderText(String cell)
  {
    ReportCell rc = new ReportCell(cell, 0);
    rc.setCellStyleName(ReportDefinition.STRING_STYLE);
    rc.setNoBorder(true);
    rc.setRow(this);
    this.cells.add(rc);
  }

  public void addCustomCell(String cell, boolean isBold, int hAlign, int vAlign, int borderWidth, boolean hasBorder, int fontSize)
  {
    ReportCell rc = new ReportCell(cell, 3);
    rc.setCellStyleName(ReportDefinition.STRING_STYLE);
    rc.setCellContent(cell);
    rc.setBold(isBold);
    rc.setHAlign(hAlign);
    rc.setVAlign(vAlign);
    rc.setBorderWidth(borderWidth);
    rc.setHasBorders(hasBorder);
    rc.setFontSize(fontSize);
    rc.setRow(this);
    this.cells.add(rc);
  }

  public void sumColumn(int colIndex, int startRow, int endRow, boolean currencyFormat)
  {
    String columnName = ColumnName(colIndex);
    String formula = "SUM(" + columnName + (startRow + 1) + ":" + columnName + (endRow + 1) + ")";
    ReportCell rc = new ReportCell(formula, -1);
    if (currencyFormat)
      rc.setCellStyleName(ReportDefinition.CURRENCY_STYLE);
    else
      rc.setCellStyleName(ReportDefinition.NUMBER_STYLE);
    this.cells.add(rc);
  }

  public void sumAbove(int startRow, int endRow, boolean currencyFormat)
  {
    int colIndex = this.cells.size() + 1;
    sumColumn(colIndex, startRow, endRow, currencyFormat);
  }

  public void addFormulaNumber(String formula)
  {
    ReportCell rc = new ReportCell(formula, -1);
    rc.setCellStyleName(ReportDefinition.NUMBER_STYLE);
    this.cells.add(rc);
  }

  public void addFormulaString(String formula)
  {
    ReportCell rc = new ReportCell(formula, -1);
    rc.setCellStyleName(ReportDefinition.STRING_STYLE);
    this.cells.add(rc);
  }

  public void setOrder(int colIndex)
  {
    this.sortKey += ((ReportCell)this.cells.get(colIndex)).getCellContent();
  }

  public List<ReportCell> getCells() {
    return this.cells;
  }

  public void setCells(List<ReportCell> cells) {
    this.cells = cells;
  }

  public int size()
  {
    return this.cells.size();
  }

  public int getRowNumber() {
    return this.rowNumber;
  }

  public void setRowNumber(int rowNumber) {
    this.rowNumber = rowNumber;
  }

  public String getSortKey() {
    return this.sortKey;
  }

  public void setSortKey(String sortKey) {
    this.sortKey = sortKey;
  }

  public static String GenerateLetter(int number)
  {
    String letter = "";
    switch (number) {
    case 0:
      letter = "Z"; break;
    case 1:
      letter = "A"; break;
    case 2:
      letter = "B"; break;
    case 3:
      letter = "C"; break;
    case 4:
      letter = "D"; break;
    case 5:
      letter = "E"; break;
    case 6:
      letter = "F"; break;
    case 7:
      letter = "G"; break;
    case 8:
      letter = "H"; break;
    case 9:
      letter = "I"; break;
    case 10:
      letter = "J"; break;
    case 11:
      letter = "K"; break;
    case 12:
      letter = "L"; break;
    case 13:
      letter = "M"; break;
    case 14:
      letter = "N"; break;
    case 15:
      letter = "O"; break;
    case 16:
      letter = "P"; break;
    case 17:
      letter = "Q"; break;
    case 18:
      letter = "R"; break;
    case 19:
      letter = "S"; break;
    case 20:
      letter = "T"; break;
    case 21:
      letter = "U"; break;
    case 22:
      letter = "V"; break;
    case 23:
      letter = "W"; break;
    case 24:
      letter = "X"; break;
    case 25:
      letter = "Y"; break;
    default:
      return "Sorry,there is no answer!";
    }
    return letter;
  }

  public static String ColumnName(int columnNum)
  {
    String columnName = "";
    int i = columnNum / 26;
    int j = columnNum % 26;
    String k = "";
    if (i == 0)
    {
      columnName = GenerateLetter(j);
    }
    else
    {
      k = GenerateLetter(j);
      if (j == 0)
      {
        if (i == 1)
        {
          return columnName = k;
        }

        return columnName = ColumnName(i - 1) + k;
      }

      columnName = ColumnName(i) + k;
    }
    return columnName;
  }

  public static void main(String[] args) {
    System.out.println(ColumnName(2));
    System.out.println(ColumnName(12));
    System.out.println(ColumnName(22));
    System.out.println(ColumnName(26));
    System.out.println(ColumnName(52));
    System.out.println(ColumnName(62));
    System.out.println(ColumnName(72));
    System.out.println(ColumnName(82));
  }
}