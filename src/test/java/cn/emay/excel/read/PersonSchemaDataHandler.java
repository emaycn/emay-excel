package cn.emay.excel.read;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import cn.emay.excel.common.Person;
import cn.emay.excel.common.schema.base.ColumnSchema;
import cn.emay.excel.common.schema.base.SheetReadSchemaParams;
import cn.emay.excel.common.schema.base.SheetSchema;
import cn.emay.excel.read.handler.SchemaSheetDataHandler;

/**
 * 
 * @author Frank
 *
 */
public class PersonSchemaDataHandler implements SchemaSheetDataHandler<Person> {

	@Override
	public SheetSchema getSheetSchema() {
		SheetReadSchemaParams sheetSchemaParams = new SheetReadSchemaParams();
		sheetSchemaParams.setReadTitleRowIndex(0);
		sheetSchemaParams.setReadColumnBy("Index");
		sheetSchemaParams.setReadDataStartRowIndex(1);
		sheetSchemaParams.setReadDataEndRowIndex(-1);
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

	private List<Person> list = new ArrayList<>();

	@Override
	public void handle(int rowIndex, Person data) {
		list.add(data);
	}

	public List<Person> getDatas() {
		return list;
	}

	@Override
	public Class<Person> getDataClass() {
		return Person.class;
	}

}
