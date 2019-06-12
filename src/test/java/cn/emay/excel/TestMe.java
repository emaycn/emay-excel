package cn.emay.excel;

import java.util.List;

import cn.emay.excel.common.schema.annotation.ExcelColumn;
import cn.emay.excel.common.schema.annotation.ExcelSheet;
import cn.emay.excel.read.ExcelReader;

/**
 * 
 * @author Frank
 *
 */
public class TestMe {

	public static void main(String[] args) {
		List<Excc> es = ExcelReader.readFirstSheet("C:\\Users\\Frank\\Desktop\\test.xls", Excc.class);
		es.stream().forEach(ex -> System.out.println(ex.toString()));
	}

	@ExcelSheet(readDataStartRowIndex = 0)
	public static class Excc {

		@ExcelColumn(index = 0, title = "年龄")
		private String f1;
		@ExcelColumn(index = 1, title = "名字")
		private String f2;
		@ExcelColumn(index = 2, title = "生日")
		private String f3;
		@ExcelColumn(index = 3, title = "创建时间")
		private String f4;
		@ExcelColumn(index = 4, title = "得分")
		private String f5;
		@ExcelColumn(index = 5, title = "是否戴眼镜")
		private String f6;
		@ExcelColumn(index = 6, title = "资产")
		private String f7;

		@Override
		public String toString() {
			return "Excc [f1=" + f1 + ", f2=" + f2 + ", f3=" + f3 + ", f4=" + f4 + ", f5=" + f5 + ", f6=" + f6 + ", f7=" + f7 + "]";
		}

	}

}
