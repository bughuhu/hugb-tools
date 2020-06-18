package cn.hugb.excel.easyExcel;

import java.util.ArrayList;
import java.util.List;

import com.alibaba.excel.EasyExcel;

public class NoTemplateWrite {

	/**
	 * 无模板写文件
	 * 
	 * @param filePath
	 * @param head     表头数据
	 * @param data     表内容数据
	 */
	public static void write(String filePath, List<List<String>> head, List<List<Object>> data) {
		EasyExcel.write(filePath).head(head).sheet().doWrite(data);
	}

	/**
	 * 无模板写文件
	 * 
	 * @param filePath
	 * @param head      表头数据
	 * @param data      表内容数据
	 * @param sheetNo   sheet页号，从0开始
	 * @param sheetName sheet名称
	 */
	public static void write(String filePath, List<List<String>> head, List<List<Object>> data, Integer sheetNo,
			String sheetName) {
		EasyExcel.write(filePath).head(head).sheet(sheetNo, sheetName).doWrite(data);
	}

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		String filePath = "D:\\noTemplate.xls";

	    List<List<String>> list = new ArrayList<List<String>>();
	    List<String> head0 = new ArrayList<String>();
	    head0.add("字符串" + System.currentTimeMillis());
	    List<String> head1 = new ArrayList<String>();
	    head1.add("数字" + System.currentTimeMillis());
	    List<String> head2 = new ArrayList<String>();
	    head2.add("日期" + System.currentTimeMillis());
	    list.add(head0);
	    list.add(head1);
	    list.add(head2);
	    
	    List<List<Object>> data = new ArrayList<List<Object>>();
	    for(int i=1;i<=100;i++) {
		    List<Object> item = new ArrayList<Object>();
		    item.add("字符串" + i);
		    item.add("字符串1" + i);
		    item.add("字符串2" + i);
		    data.add(item);
	    }


	    NoTemplateWrite.write(filePath, list, data);
	}

}
