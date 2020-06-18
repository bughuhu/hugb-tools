package cn.hugb.easyExcel;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import cn.hutool.core.convert.Convert;
import lombok.Data;

/**
 * 
 * @author huguobiao
 *
 */
@Data
public class ReportRow {

	List<Object> cells = new ArrayList<Object>();

	public void addString(String value) {
		this.cells.add(value);
	}
	
	public void addNumber(String cell) {
        this.cells.add(Double.valueOf(cell));
	}

	public void addInteger(String cell) {
        this.cells.add(Integer.valueOf(cell));
	}

	public void addDecimal(String cell) {
        this.cells.add(Convert.toBigDecimal(cell));
	}
	
	public void addDate(String cell) {
        try {
        	SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        	Date date =  sdf.parse(cell);
			this.cells.add(date);
		} catch (ParseException e) {
			e.printStackTrace();
		}
	}

	public void addDateTime(String cell) {
        try {
			this.cells.add(new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").parse(cell));
		} catch (ParseException e) {
			e.printStackTrace();
		}
	}
	
	public void addCell(String cell) {
		addString(cell);
	}

}
