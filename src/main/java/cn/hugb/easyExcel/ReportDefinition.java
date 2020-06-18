package cn.hugb.easyExcel;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import com.alibaba.excel.write.style.column.LongestMatchColumnWidthStyleStrategy;

import lombok.Data;

/**
 * 
 * @author huguobiao
 *
 */
@Data
public class ReportDefinition {

	private String sheetName = "sheet1";
	private String reportName = "";
	private String operatorName = "";

	// private List<ReportParameter> parametes = new ArrayList<ReportParameter>();
	private List<List<String>> headers = new ArrayList<List<String>>();
	private List<List<Object>> rows = new ArrayList<List<Object>>();

	/**
	 * 添加数据行
	 * 
	 * @param row
	 */
	public void addRow(ReportRow row) {
		this.rows.add(row.getCells());
	}

	/**
	 * 添加头行
	 * 
	 * @param row
	 */
	public void addHeaderRow(ReportRow headerRow) {
		List<Object> dataList = headerRow.getCells();
		for (int i = 0; i < dataList.size(); i++) {
			List<String> item = new ArrayList<String>();
			item.add(dataList.get(i).toString());
			this.headers.add(item);
		}
	}

	public void buildExcel(String filePath) {
		try {

			// 头的策略
			WriteCellStyle headWriteCellStyle = new WriteCellStyle();

			// 背景设置为红色
			headWriteCellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
			WriteFont headWriteFont = new WriteFont();
			headWriteFont.setFontHeightInPoints((short) 12);
			headWriteCellStyle.setWriteFont(headWriteFont);
			headWriteCellStyle.setBorderBottom(BorderStyle.THIN);
			headWriteCellStyle.setBorderLeft(BorderStyle.THIN);
			headWriteCellStyle.setBorderRight(BorderStyle.THIN);
			headWriteCellStyle.setBorderTop(BorderStyle.THIN);

			// 内容的策略
			WriteCellStyle contentWriteCellStyle = new WriteCellStyle();

			// 这里需要指定 FillPatternType 为FillPatternType.SOLID_FOREGROUND 不然无法显示背景颜色.头默认了
			// FillPatternType所以可以不指定
			contentWriteCellStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);

			// 背景绿色
			contentWriteCellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
			contentWriteCellStyle.setBorderBottom(BorderStyle.THIN);
			contentWriteCellStyle.setBorderLeft(BorderStyle.THIN);
			contentWriteCellStyle.setBorderRight(BorderStyle.THIN);
			contentWriteCellStyle.setBorderTop(BorderStyle.THIN);

			WriteFont contentWriteFont = new WriteFont();
			// 字体大小
			contentWriteFont.setFontHeightInPoints((short) 10);
			contentWriteCellStyle.setWriteFont(contentWriteFont);

			// 这个策略是 头是头的样式 内容是内容的样式 其他的策略可以自己实现
			HorizontalCellStyleStrategy horizontalCellStyleStrategy = new HorizontalCellStyleStrategy(
					headWriteCellStyle, contentWriteCellStyle);

			// 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
			EasyExcel.write(filePath).registerWriteHandler(new LongestMatchColumnWidthStyleStrategy())
					.registerWriteHandler(horizontalCellStyleStrategy).head(headers).sheet(sheetName).doWrite(rows);

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void main(String[] args) {

		ReportDefinition definition = new ReportDefinition();
		definition.setOperatorName("hugb");
		definition.setReportName("报表名称");

		ReportRow headerRow = new ReportRow();
		headerRow.addCell("子公司");
		headerRow.addCell("一级部门");
		headerRow.addCell("二级部门");
		headerRow.addCell("三级部门");
		headerRow.addCell("四级部门");
		headerRow.addCell("正式员工");
		headerRow.addCell("正式员工调整");
		headerRow.addCell("正式员工小计");
		headerRow.addCell("派遣\n员工");
		headerRow.addCell("派遣员工\n调整");
		headerRow.addCell("派遣员工\n小计");
		headerRow.addCell("总计");
		headerRow.addCell("正式\n员工");
		headerRow.addCell("正式员工\n调整");
		headerRow.addCell("正式员工\n小计");
		headerRow.addCell("派遣员工");
		headerRow.addCell("派遣员工\n调整");
		headerRow.addCell("派遣员工小计");
		headerRow.addCell("总计");
		definition.addHeaderRow(headerRow);

		ReportRow headerRow2 = new ReportRow();
		headerRow2.addCell("子公司");
		headerRow2.addCell("一级部门");
		headerRow2.addCell("二级部门");
		headerRow2.addCell("三级部门");
		headerRow2.addCell("四级部门");
		headerRow2.addCell("正式员工");
		headerRow2.addCell("正式员工调整");
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

		definition.addRow(headerRow2);

		for (int i = 1; i < 1000000; i++) {

			ReportRow dataRow3 = new ReportRow();
			dataRow3.addCell("总部");
			dataRow3.addCell("技术部");
			dataRow3.addCell("");
			dataRow3.addCell("");
			dataRow3.addDate("2012-03-01");
			dataRow3.addCell("1");
			dataRow3.addCell("2");
			dataRow3.addCell("3");
			dataRow3.addCell("3");
			dataRow3.addCell("4");
			dataRow3.addCell("1111");
			dataRow3.addNumber("15001230.23");
			dataRow3.addCell("0");
			dataRow3.addInteger("150");
			dataRow3.addCell("150");
			dataRow3.addDateTime("2020-06-18 23:59:55");
			dataRow3.addCell("25.0");
			dataRow3.addDecimal("150");
			definition.addRow(dataRow3);
		}

		definition.buildExcel("D:\\sample.xlsx");
		System.out.println("------The End-------");
	}

}
