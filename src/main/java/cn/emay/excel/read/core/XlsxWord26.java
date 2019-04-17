package cn.emay.excel.read.core;

/**
 * 大写字母26进制
 * 
 * @author Frank
 *
 */
public class XlsxWord26 {

	/**
	 * 进制数
	 */
	private static long HEX;

	/**
	 * 进制字符集
	 */
	private static char[] CHARS;

	/**
	 * 进制对应的数字集
	 */
	private static Integer[] INTS;

	/**
	 * 初始化
	 * 
	 */
	static {
		CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".toCharArray();
		HEX = CHARS.length;
		INTS = new Integer[CHARS.length];
		for (int i = 0; i < HEX; i++) {
			INTS[i] = i;
		}
	}

	/**
	 * 根据字符序号拿到字符
	 * 
	 * @param index
	 * @return
	 */
	private static Character getCharIndex(int index) {
		return CHARS[index];
	}

	/**
	 * 根据字符拿到数字
	 * 
	 * @param cha
	 * @return
	 */
	private static Integer getIndexChar(char cha) {
		int index = -1;
		for (int i = 0; i < HEX; i++) {
			if (CHARS[i] == cha) {
				index = i;
				break;
			}
		}
		if (index == -1) {
			return null;
		}
		return INTS[index];
	}

	/**
	 * 编码
	 * 
	 * @param number
	 * @return
	 */
	public static String encode(long number) {
		if (number <= 0) {
			return new String(new char[] { CHARS[0] });
		}
		Long rest = number;
		char[] stac = new char[64];
		int i = 64;
		while (rest != 0) {
			Long res = rest / HEX;
			int env = new Long(rest % HEX).intValue();
			Character cha = getCharIndex(env);
			stac[--i] = cha;
			rest = res;
		}
		return new String(stac, i, (64 - i));
	}

	/**
	 * 解码
	 * 
	 * @param inhex
	 * @return
	 */
	public static long decode(String inhex) {
		if (inhex == null || inhex.trim().length() == 0) {
			return 0L;
		}
		long result = 0L;
		char[] chars = inhex.toCharArray();
		for (int i = 0; i < chars.length; i++) {
			char cha = chars[i];
			Integer ind = getIndexChar(cha);
			if (ind == null) {
				return 0L;
			}
			result += Math.pow(HEX, chars.length - i - 1) * ind;
		}
		return result;
	}

}