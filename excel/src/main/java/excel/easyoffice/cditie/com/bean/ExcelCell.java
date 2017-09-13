package excel.easyoffice.cditie.com.bean;

import lombok.Data;

@Data
public class ExcelCell implements PoiCell{
	/**列名**/
	private String columnName;
	
	/**数据key**/
	private String datakey;
	
	
	

	@Override
	public boolean isMerge() {
		return false;
	}
	@Override
	public String getValue() {
		return columnName;
	}
	public ExcelCell(String columnName, String datakey) {
		super();
		this.columnName = columnName;
		this.datakey = datakey;
	}
	@Override
	public String getKey() {
		// TODO Auto-generated method stub
		return datakey;
	}
	
}
