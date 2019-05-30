package cn.emay.excel;

import java.util.List;

import cn.emay.excel.common.schema.annotation.ExcelColumn;
import cn.emay.excel.common.schema.annotation.ExcelSheet;
import cn.emay.excel.read.ExcelReader;

public class TestMe {
	
	public static void main(String[] args) {
		List<Excc> es = ExcelReader.readFirstSheet("C:\\Users\\Frank\\Desktop\\test.xls", Excc.class);
		es.stream().forEach(ex -> System.out.println(ex.getMobile()));
	}
	
	@ExcelSheet(readDataStartRowIndex = 0)
	public static class Excc{
		
		@ExcelColumn(index = 0, title = "手机")
		private String mobile;

		public String getMobile() {
			return mobile;
		}

		public void setMobile(String mobile) {
			this.mobile = mobile;
		}
		
	}


}


