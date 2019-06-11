package cn.emay.excel;

import java.util.List;

import cn.emay.excel.common.schema.annotation.ExcelColumn;
import cn.emay.excel.common.schema.annotation.ExcelSheet;
import cn.emay.excel.read.ExcelReader;

public class TestMe {

	public static void main(String[] args) {
		List<Excc> es = ExcelReader.readFirstSheet("C:\\Users\\Frank\\Desktop\\test.xlsx", Excc.class);
		es.stream().forEach(ex -> System.out.println(ex.toString()));
	}

	@ExcelSheet(readDataStartRowIndex = 1)
	public static class Excc {

		@ExcelColumn(index = 0, title = "手机")
		private String mobile;
		@ExcelColumn(index = 1, title = "用户")
		private String user;
		@ExcelColumn(index = 2, title = "编号")
		private String number;
		@ExcelColumn(index = 3, title = "余额")
		private String mount;

		public String getMobile() {
			return mobile;
		}

		public void setMobile(String mobile) {
			this.mobile = mobile;
		}

		public String getUser() {
			return user;
		}

		public void setUser(String user) {
			this.user = user;
		}

		public String getNumber() {
			return number;
		}

		public void setNumber(String number) {
			this.number = number;
		}

		public String getMount() {
			return mount;
		}

		public void setMount(String mount) {
			this.mount = mount;
		}

		@Override
		public String toString() {
			return "Excc [mobile=" + mobile + ", user=" + user + ", number=" + number + ", mount=" + mount + "]";
		}

	}

}
