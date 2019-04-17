# 亿美JAVA组件：Excel组件


**OO思想，基于注解的快速读取Excel组件，有效降低编码复杂度。**

## 1. 实例代码

### 1.1 Schema定义

```java
@ExcelSheet
class User{

	@ExcelColumn(columnIndex = 0, title = "用户名")
	String userName;
	@ExcelColumn(columnIndex = 1, title = "密码")
	String passWord;

}
```

### 1.2 写Excel

```java

// schema方式写入，小数据量使用
List<Person> list = new ArrayList<Person>();
ExcelWriter.writeExcel(excelPath, list);

// schema方式写入，大数据量使用
final List<Person> list = new ArrayList<>();
ExcelWriter.writeExcel(xlsxPath, new  DataWriter<Person>() {
	
	@Override
	public Person getData(int rowIndex) {
		return list.get(rowIndex);
	}

	@Override
	public boolean hasData(int rowIndex) {
		return rowIndex < list.size();
	}

	@Override
	public Class<Person> getShemaClass() {
		return Person.class;
	}
});
```

### 1.3 读Excel

```java
// schema方式读取，有返回值，小数据量使用
List<Person> list = ExcelReader.readFirstSheetWithSchema(excelPath, Person.class);

// schema方式读取，无返回值，大数据量使用
ExcelReader.readFirstSheetWithReader(excelPath, new DataReader<Person>() {

	@Override
	public void handlerRowData(int rowIndex, Person data) {
		System.out.println(data.toString());
	}

	@Override
	public Class<Person> getShemaClass() {
		return Person.class;
	}
});
```

### 1.4 更多支持

**更丰富的支持，请参照src/test/java/cn.emay.excel.ExcelTest.java**

## 2 注意事项

1. 如果需要给单元格设置样式，直接用cell.getCellStyle();
2. Excel支持的数字最大为15位，超过15位会失去精度，存储时请注意以文本格式存储;
3. 如果需要单元格颜色，xls中的GREY_25_PERCENT、GREY_40_PERCENT两个颜色将被替换;
4. schema方式仅支持String,Long,Integer,Double,Boolean,Date,BigDecimal类型;
5. 写入Excel的null值，读取String类型，会读出"";
