package cn.hugb.validate;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class EmailValidate {

	/**
	 * 判断是否是正确的邮件格式
	 * 
	 * @param string
	 * @return
	 */
	public static boolean isEmail(String string) {
		if (string == null)
			return false;
		String regEx1 = "^([a-z0-9A-Z]+[-|\\.]?)+[a-z0-9A-Z]@([a-z0-9A-Z]+(-[a-z0-9A-Z]+)?\\.)+[a-zA-Z]{2,}$";
		Pattern p;
		Matcher m;
		p = Pattern.compile(regEx1);
		m = p.matcher(string);
		if (m.matches())
			return true;
		else
			return false;
	}

	public static void main(String[] args) {
		System.out.println("softworm@126.com: " + isEmail("softworm@126.com"));
		System.out.println("softworm126.com: " + isEmail("softworm126.com"));
	}

}
