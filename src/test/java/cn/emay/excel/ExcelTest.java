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
import cn.emay.excel.read.NormalReader;
import cn.emay.excel.read.PersonDataHandler;
import cn.emay.excel.read.PersonSchemaDataHandler;
import cn.emay.excel.write.ExcelWriter;
import cn.emay.excel.write.NormalWriter;
import cn.emay.excel.write.PersonDataGetter;
import cn.emay.excel.write.PersonSchemaDataGetter;
import cn.emay.excel.write.data.WriteData;

/**
 * 5组测试:<br/>
 * 1.normal--基于writer/reader的测试<br/>
 * 2.schema--基于自定义writer/reader的schemaHandler测试<br/> 
 * 3.annscheman--基于注解的writer/reader的Handler测试<br/>
 * 4.data--基于注解的writer/reader的数据/class测试<br/>
 * 5.coordinate--基于坐标的精准读取测试<br/>
 * 
 * @author Frank
 *
 */
@FixMethodOrder(MethodSorters.NAME_ASCENDING)
public class ExcelTest {

	public static String xlsPath;
	public static String xlsxPath;
	public static String xlsPathTo;
	public static String xlsxPathTo;

	private static List<Person> datas = new ArrayList<Person>(7);
	private static List<String> titles = new ArrayList<String>(7);

	private static WriteData[] writedDatas = new WriteData[8];

	private static int[][] coordinates = new int[8][];

	@BeforeClass
	public static void before() {
		String dirPath = System.getProperty("user.dir") + File.separator;
		xlsPath = dirPath + File.separatorChar + "exceltest.xls";
		xlsxPath = dirPath + File.separatorChar + "exceltest.xlsx";
		xlsPathTo = dirPath + File.separatorChar + "exceltest.to.xls";
		xlsxPathTo = dirPath + File.separatorChar + "exceltest.to.xlsx";
		new File(xlsPath).delete();
		new File(xlsxPath).delete();
		new File(xlsPathTo).delete();
		new File(xlsxPathTo).delete();

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

		writedDatas[0] = new WriteData(0, 0, 1, "001", null);
		writedDatas[1] = new WriteData(0, 1, 2, "012", null);
		writedDatas[2] = new WriteData(0, 2, 3, "023", null);
		writedDatas[3] = new WriteData(0, 3, 4, "034", null);
		writedDatas[4] = new WriteData(0, 5, 5, "055", null);
		writedDatas[5] = new WriteData(1, 0, 1, "101", null);
		writedDatas[6] = new WriteData(1, 4, 3, "143", null);
		writedDatas[7] = new WriteData(2, 0, 1, "201", null);

		for (int i = 0; i < writedDatas.length; i++) {
			coordinates[i] = writedDatas[i].getCoordinate();
		}
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
		ExcelWriter.write(xlsPath, new NormalWriter(titles, datas));
		NormalReader reader = new NormalReader();
		ExcelReader.readFirstSheet(xlsPath, reader);
		List<Person> list = reader.getDatas();
		List<String> tits = reader.getTitles();
		checkTitle(tits);
		check(list);
	}

	@Test
	public void normalXlsxTest() {
		ExcelWriter.write(xlsxPath, new NormalWriter(titles, datas));
		NormalReader reader = new NormalReader();
		ExcelReader.readFirstSheet(xlsxPath, reader);
		List<Person> list = reader.getDatas();
		List<String> tits = reader.getTitles();
		checkTitle(tits);
		check(list);
	}

	@Test
	public void annSchemaXlsTest() {
		ExcelWriter.write(xlsPath, new PersonDataGetter(datas));
		PersonDataHandler red = new PersonDataHandler();
		ExcelReader.readFirstSheet(xlsPath, red);
		List<Person> list = red.getDatas();
		check(list);
	}

	@Test
	public void annSchemaXlsxTest() {
		ExcelWriter.write(xlsxPath, new PersonDataGetter(datas));
		PersonDataHandler red = new PersonDataHandler();
		ExcelReader.readFirstSheet(xlsxPath, red);
		List<Person> list = red.getDatas();
		check(list);
	}
	
	@Test
	public void schemaXlsTest() {
		ExcelWriter.write(xlsPath, new PersonSchemaDataGetter(datas));
		PersonSchemaDataHandler red = new PersonSchemaDataHandler();
		ExcelReader.readFirstSheet(xlsPath, red);
		List<Person> list = red.getDatas();
		check(list);
	}

	@Test
	public void schemaXlsxTest() {
		ExcelWriter.write(xlsxPath, new PersonSchemaDataGetter(datas));
		PersonSchemaDataHandler red = new PersonSchemaDataHandler();
		ExcelReader.readFirstSheet(xlsxPath, red);
		List<Person> list = red.getDatas();
		check(list);
	}

	@Test
	public void dataXlsTest() {
		ExcelWriter.write(xlsPath, datas);
		List<Person> list = ExcelReader.readFirstSheet(xlsPath, Person.class);
		check(list);
	}

	@Test
	public void dataXlsxTest() {
		ExcelWriter.write(xlsxPath, datas);
		List<Person> list = ExcelReader.readFirstSheet(xlsxPath, Person.class);
		check(list);
	}

	@Test
	public void coordinateXlsxTest() {
		ExcelWriter.write(xlsxPath, new NormalWriter(titles, datas));
		ExcelWriter.writeByCoordinate(xlsxPath, xlsxPathTo, writedDatas);
		List<String> resus = ExcelReader.readByCoordinate(String.class,xlsxPathTo, coordinates);
		for (int i = 0; i < 8; i++) {
			Assert.assertEquals(resus.get(i), writedDatas[i].getData());
		}
	}

	@Test
	public void coordinateXlsTest() {
		ExcelWriter.write(xlsPath, new NormalWriter(titles, datas));
		ExcelWriter.writeByCoordinate(xlsPath, xlsPathTo, writedDatas);
		List<String> resus = ExcelReader.readByCoordinate(String.class,xlsPathTo, coordinates);
		for (int i = 0; i < 8; i++) {
			Assert.assertEquals(resus.get(i), writedDatas[i].getData());
		}
	}

	@After
	public void after() {
		new File(xlsPath).delete();
		new File(xlsxPath).delete();
		new File(xlsPathTo).delete();
		new File(xlsxPathTo).delete();
	}

}
