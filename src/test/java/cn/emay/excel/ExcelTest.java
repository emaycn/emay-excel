package cn.emay.excel;

import java.io.File;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.junit.After;
import org.junit.Assert;
import org.junit.BeforeClass;
import org.junit.FixMethodOrder;
import org.junit.Test;
import org.junit.runners.MethodSorters;

import cn.emay.excel.common.Person;
import cn.emay.excel.read.ExcelReader;
import cn.emay.excel.reader.handler.PersonDataReader;
import cn.emay.excel.reader.handler.ReadNormalHandler;
import cn.emay.excel.write.ExcelWriter;
import cn.emay.excel.writer.handler.PersonDataWriter;
import cn.emay.excel.writer.handler.WriteNormalHandler;

@FixMethodOrder(MethodSorters.NAME_ASCENDING)
public class ExcelTest {

	public static String xlsPath;
	public static String xlsxPath;

	private static List<Person> datas = new ArrayList<Person>(7);
	private static List<String> titles = new ArrayList<String>(7);

	 @BeforeClass
	public static void before() {
		String dirPath = System.getProperty("user.dir") + File.separator;
		xlsPath = dirPath + File.separatorChar + "exceltest.xls";
		xlsxPath = dirPath + File.separatorChar + "exceltest.xlsx";
		new File(xlsPath).delete();
		new File(xlsxPath).delete();

		titles.add("年龄");
		titles.add("名字");
		titles.add("生日");
		titles.add("创建时间");
		titles.add("得分");
		titles.add("是否戴眼镜");
		titles.add("资产");

		datas.add(new Person(-1, "李四lili", new Date(System.currentTimeMillis() - 1000L * 60L), 12345678901234L, 37.991, false, new BigDecimal(123.123123d)));
		datas.add(new Person(Integer.MAX_VALUE, null, new Date(System.currentTimeMillis() / 10000L), 2345678901234L, 1.0, false, new BigDecimal("99.99123")));
		// Excel支持的数字最大为15位，超过15位会失去精度，所以存储时要转成文本存储。
		datas.add(new Person(79, "" + Long.MAX_VALUE, new Date(System.currentTimeMillis() - 1000L * 6000L), 345678901234L, 2.3332, true, new BigDecimal(23444.3123d)));
	}

	private void check(List<Person> list) {
		for (int i = 0; i < datas.size(); i++) {
			Person old = datas.get(i);
			Person ne = list.get(i);
			Assert.assertEquals(old.getCreateTime(), ne.getCreateTime());
			// 写入null，会读出""
			Assert.assertEquals(old.getName(), "".equals(ne.getName()) ? null : ne.getName());
			Assert.assertEquals(old.getAge(), ne.getAge());
			Assert.assertEquals(old.getBrith(), ne.getBrith());
			Assert.assertEquals(old.getHasGlass(), ne.getHasGlass());
			Assert.assertEquals(old.getMoney().setScale(4, BigDecimal.ROUND_HALF_UP).toString(), ne.getMoney().setScale(4, BigDecimal.ROUND_HALF_UP).toString());
			Assert.assertEquals(new BigDecimal(old.getScore()).setScale(2, BigDecimal.ROUND_HALF_UP).toString(), new BigDecimal(old.getScore()).setScale(2, BigDecimal.ROUND_HALF_UP).toString());
		}
	}

	private void checkTitle(List<String> tits) {
		for (int i = 0; i < titles.size(); i++) {
			String old = titles.get(i);
			String ne = tits.get(i);
			Assert.assertEquals(old, ne);
		}
	}

	 @Test
	public void normalXlsTest() {
		ExcelWriter.write(xlsPath, new WriteNormalHandler(titles, datas));
		ReadNormalHandler handler = new ReadNormalHandler();
		ExcelReader.readFirstSheet(xlsPath, handler);
		List<Person> list = handler.getDatas();
		List<String> tits = handler.getTitles();
		checkTitle(tits);
		check(list);
	}

	 @Test
	public void normalXlsxTest() {
		ExcelWriter.write(xlsxPath, new WriteNormalHandler(titles, datas));
		ReadNormalHandler handler = new ReadNormalHandler();
		ExcelReader.readFirstSheet(xlsxPath, handler);
		List<Person> list = handler.getDatas();
		List<String> tits = handler.getTitles();
		checkTitle(tits);
		check(list);
	}

	 @Test
	public void dataXlsTest() {
		ExcelWriter.write(xlsPath, new PersonDataWriter(datas));
		PersonDataReader red = new PersonDataReader();
		ExcelReader.readFirstSheet(xlsPath, red);
		List<Person> list = red.getDatas();
		check(list);
	}

	 @Test
	public void dataXlsxTest() {
		ExcelWriter.write(xlsxPath, new PersonDataWriter(datas));
		PersonDataReader red = new PersonDataReader();
		ExcelReader.readFirstSheet(xlsxPath, red);
		List<Person> list = red.getDatas();
		check(list);
	}

	 @Test
	public void schemaXlsTest() {
		ExcelWriter.write(xlsPath, datas);
		List<Person> list = ExcelReader.readFirstSheet(xlsPath, Person.class);
		check(list);
	}

	 @Test
	public void schemaTest() {
		ExcelWriter.write(xlsxPath, datas);
		List<Person> list = ExcelReader.readFirstSheet(xlsxPath, Person.class);
		check(list);
	}

	 @After
	public void after() {
		new File(xlsPath).delete();
		new File(xlsxPath).delete();
	}

}
