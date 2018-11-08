package cn.hugb.util;

/**
 * 通过用工具类
 */
public class CommonKit {

	public static final String[] BAI_JIA_XING_DOUBLE = { "欧阳", "太史", "上官", "端木", "司马", "东方", "独孤", "南宫", "万俟", "闻人",
			"夏侯", "诸葛", "尉迟", "公羊", "赫连", "澹台", "皇甫", "宗政", "濮阳", "公冶", "太叔", "申屠", "公孙", "慕容", "仲孙", "钟离", "长孙", "宇文",
			"司徒", "鲜于", "司空", "闾丘", "子车", "亓官", "司寇", "巫马", "公西", "颛孙", "壤驷", "公良", "漆雕", "乐正", "宰父", "谷梁", "拓跋", "夹谷",
			"轩辕", "令狐", "段干", "百里", "呼延", "东郭", "南门", "羊舌", "微生", "公户", "公玉", "公仪", "梁丘", "公仲", "公上", "公门", "公山", "公坚",
			"左丘", "公伯", "西门", "公祖", "第五", "公乘", "贯丘", "公皙", "南荣", "东里", "东宫", "仲长", "子书", "子桑", "即墨", "达奚", "褚师",
			"吴铭" };

	/**
	 * 校验是否是复姓
	 */
	public static boolean checkName(String name) {
		String lastName = name.substring(0, 2);
		for (int i = 0; i < BAI_JIA_XING_DOUBLE.length; i++) {
			if (lastName.equals(BAI_JIA_XING_DOUBLE[i])) {
				return true;
			}
		}
		return false;
	}

	/**
	 * 获取姓
	 */
	public static String getLastName(String name) {
		if (checkName(name)) {
			return name.substring(0, 2);
		} else {
			return name.substring(0, 1);
		}
	}

	/**
	 * 获取名
	 */
	public static String getFirstName(String name) {
		if (checkName(name)) {
			return name.substring(2);
		} else {
			return name.substring(1);
		}
	}

	public static void main(String[] args) {
		System.out.println("胡国彪: 姓=" + getLastName("胡国彪") + ",名=" + getFirstName("胡国彪"));
		System.out.println("上官无汲: 姓=" + getLastName("上官无汲") + ",名=" + getFirstName("上官无汲"));
		System.out.println("慕容复: 姓=" + getLastName("慕容复") + ",名=" + getFirstName("慕容复"));
		System.out.println("王菲: 姓=" + getLastName("王菲") + ",名=" + getFirstName("王菲"));
	}

}
