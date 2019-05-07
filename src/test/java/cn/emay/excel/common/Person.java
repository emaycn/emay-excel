package cn.emay.excel.common;

import java.math.BigDecimal;
import java.util.Date;

import cn.emay.excel.common.schema.annotation.ExcelColumn;
import cn.emay.excel.common.schema.annotation.ExcelSheet;

/**
 * 人
 * 
 * @author Frank
 *
 */
@ExcelSheet(isNeedBorder = true)
public class Person {

	/**
	 * 年龄
	 */
	@ExcelColumn(index = 0, title = "年龄")
	private Integer age;
	/**
	 * 名字
	 */
	@ExcelColumn(index = 1, title = "名字")
	private String name;
	/**
	 * 生日
	 */
	@ExcelColumn(index = 2, express = "yyyy-MM-dd HH:mm:ss", title = "生日")
	private Date brith;
	/**
	 * 创建时间
	 */
	@ExcelColumn(index = 3, title = "创建时间")
	private long createTime;
	/**
	 * 得分
	 */
	@ExcelColumn(index = 4, express = "2", title = "得分")
	private Double score;
	/**
	 * 是否戴眼镜
	 */
	@ExcelColumn(index = 5, title = "是否戴眼镜")
	private Boolean hasGlass;

	/**
	 * 资产
	 */
	@ExcelColumn(index = 6, express = "4", title = "资产")
	private BigDecimal money;

	public Person() {

	}

	public Person(Integer age, String name, Date brith, Long createTime, Double score, Boolean hasGlass, BigDecimal money) {
		super();
		this.age = age;
		this.name = name;
		this.brith = brith;
		this.createTime = createTime;
		this.score = score;
		this.hasGlass = hasGlass;
		this.money = money;
	}

	@Override
	public String toString() {
		return "Person [age=" + age + ", name=" + name + ", brith=" + brith + ", createTime=" + createTime + ", score=" + score + ", hasGlass=" + hasGlass + ", money=" + money + "]";
	}

	public Integer getAge() {
		return age;
	}

	public void setAge(Integer age) {
		this.age = age;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public Date getBrith() {
		return brith;
	}

	public void setBrith(Date brith) {
		this.brith = brith;
	}

	public long getCreateTime() {
		return createTime;
	}

	public Double getScore() {
		return score;
	}

	public void setScore(Double score) {
		this.score = score;
	}

	public Boolean getHasGlass() {
		return hasGlass;
	}

	public void setHasGlass(Boolean hasGlass) {
		this.hasGlass = hasGlass;
	}

	public BigDecimal getMoney() {
		return money;
	}

	public void setMoney(BigDecimal money) {
		this.money = money;
	}

	public void setCreateTime(long createTime) {
		this.createTime = createTime;
	}

}
