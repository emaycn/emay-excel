package cn.emay.excel.write;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import cn.emay.excel.common.Person;
import cn.emay.excel.common.schema.base.ColumnSchema;
import cn.emay.excel.common.schema.base.SheetSchema;
import cn.emay.excel.common.schema.base.SheetWriteSchemaParams;
import cn.emay.excel.write.data.SchemaSheetDataGetter;

/**
 * 
 * @author Frank
 *
 */
public class PersonSchemaDataGetter implements SchemaSheetDataGetter<Person> {

	private List<Person> datas;
	int size;

	public PersonSchemaDataGetter(List<Person> datas) {
		this.datas = datas;
		size = datas.size();
	}

	@Override
	public SheetSchema getSheetSchema() {
		SheetWriteSchemaParams sheetSchemaParams = new SheetWriteSchemaParams();
		sheetSchemaParams.setAutoWidth(true);
		sheetSchemaParams.setAutoWrap(true);
		sheetSchemaParams.setCacheNumber(1000);
		sheetSchemaParams.setContentRgbColor(new int[] { 220, 230, 241 });
		sheetSchemaParams.setNeedBorder(true);
		sheetSchemaParams.setTitleRgbColor(new int[] { 250, 191, 143 });
		sheetSchemaParams.setWriteSheetName("person");
		sheetSchemaParams.setWriteTile(true);
		Map<String, ColumnSchema> columnSchemaByFieldNames = new HashMap<String, ColumnSchema>(8);
		columnSchemaByFieldNames.put("age", new ColumnSchema(0, "年龄", null));
		columnSchemaByFieldNames.put("name", new ColumnSchema(1, "名字", null));
		columnSchemaByFieldNames.put("brith", new ColumnSchema(2, "生日", "yyyy-MM-dd HH:mm:ss"));
		columnSchemaByFieldNames.put("createTime", new ColumnSchema(3, "创建时间", null));
		columnSchemaByFieldNames.put("score", new ColumnSchema(4, "得分", "2"));
		columnSchemaByFieldNames.put("hasGlass", new ColumnSchema(5, "是否戴眼镜", null));
		columnSchemaByFieldNames.put("money", new ColumnSchema(6, "资产", "4"));
		return new SheetSchema(sheetSchemaParams, columnSchemaByFieldNames);
	}

	@Override
	public Person getData(int rowIndex) {
		return datas.get(rowIndex);
	}

	@Override
	public boolean hasData(int rowIndex) {
		return rowIndex < size;
	}

	@Override
	public Class<Person> getDataClass() {
		return Person.class;
	}

}
