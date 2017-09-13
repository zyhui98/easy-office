package com.cditie.easyoffice.excel.bean;

public interface PoiCell {
	/**是否合并单元格**/
	public boolean isMerge();
	
	public String getValue();
	
	public String getKey();
}
