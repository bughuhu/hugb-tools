package cn.hugb.excel;

import java.util.Comparator;

public class ReportRowComparator implements Comparator<ReportRow> {
	public int compare(ReportRow o1, ReportRow o2) {
		return o1.getSortKey().compareTo(o2.getSortKey());
	}
}