package utils;



/**
 * @author LiMing
 *
 */
public class ExcelEntity {
	
	public ExcelEntity() {
	}
	
	public ExcelEntity(String id, String name) {
		this.id = id;
		this.name = name;
	}
	//描述改属性在excel中第0列，列名为  序号
	@MyAnnotation(columnIndex=0,columnName="序号")
	private String id;
	
	//描述改属性在excel中第1列，列名为 名字
	@MyAnnotation(columnIndex=1,columnName="名字")
	private String name;
	
	
	public String getId() {
		return id;
	}
	public void setId(String id) {
		this.id = id;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	@Override
	public String toString() {
		return "ExcelEntity [id=" + id + ", name=" + name + "]";
	}
	
	
	
	
}
