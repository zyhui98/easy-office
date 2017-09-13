package excel.easyoffice.cditie.com.bean;

import java.util.List;


import lombok.Data;
import org.springframework.util.StringUtils;

@Data
public class ExcelMergeCell implements PoiCell{
	
	/**列名**/
	private String columnName;
	
	/**数据key**/
	private String datakey;
	
	private List<PoiCell> subCell;
	
	private int mergeSize = 0;
	
	public ExcelMergeCell(String columnName, String datakey) {
		super();
		this.columnName = columnName;
		this.datakey = datakey;
	}

	public ExcelMergeCell(String columnName, String datakey, List<PoiCell> subCell, int mergeSize) {
		super();
		this.columnName = columnName;
		this.datakey = datakey;
		this.subCell = subCell;
		this.mergeSize = mergeSize;
	}
	public ExcelMergeCell(String columnName, String datakey, List<PoiCell> subCell) {
		super();
		this.columnName = columnName;
		this.datakey = datakey;
		this.subCell = subCell;
		if(!StringUtils.isEmpty(subCell)){
			this.mergeSize = subCell.size();
		}
	}

	

	@Override
	public boolean isMerge() {
		return true;
	}

	@Override
	public String getValue() {
		return columnName;
	}

	@Override
	public String getKey() {
		// TODO Auto-generated method stub
		return datakey;
	}



	
	
	
}