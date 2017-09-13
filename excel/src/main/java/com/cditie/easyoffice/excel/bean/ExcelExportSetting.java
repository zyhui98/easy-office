package com.cditie.easyoffice.excel.bean;

import java.util.List;
import java.util.Map;

import lombok.Data;

/**
 * excel配置参数类
 * @author jonny
 * @date 2016年12月9日
 */
@Data
public class ExcelExportSetting {

	private List<PoiCell> headerRow;
	
	/**数据集合list**/
	private List<Map<String, Object>> dataList;
	
	/**列头占用的行**/
	private int shiftRow = 0;
	
	/**留空行**/
	private int blankRow = 0;
	
	/**合并单元格其实列**/
	private int mergeStartIndex = 0;
	
	
	/**自定义数据**/
	private Map<String, Object> definedMap;
	

}



