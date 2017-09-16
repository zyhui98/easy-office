# Introduction
- office 操作工具包
- excel 添加导入导出
- 支持excel合并单元格、模板设置样式,jexcel和poi操作同时支持。


# Quick Start
- 参考web项目中的类
- com.cditie.easyoffice.web.controller.ExcelController;

```java

//设置excel模板
Map<String, Object> templateParams = Maps.newHashMap();
XLSTransformer transformer = new XLSTransformer();
wb = transformer.transformXLS(App.class.getResourceAsStream("/xls/excel.xlsx"), templateParams);
Sheet billInfoSheet = wb.getSheet("sheet1");

//设置excel展示配置
ExcelExportSetting excelExportSetting = new ExcelExportSetting();
List<PoiCell> cellList = Lists.newArrayList();
//一行数据的第一列
cellList.add(new ExcelMergeCell("日期","date"));
cellList.add(new ExcelMergeCell("日期1","date1"));

//一行数据的第二个列合并单元格的
ExcelMergeCell excelMergeCell = new ExcelMergeCell("自动电核笔数","zidonghebishu",
		Arrays.asList(new ExcelCell("大学贷","daxuedai"),
				new ExcelCell("手机贷","shoujidai"),
				new ExcelCell("自然贷","zirandai")));
cellList.add(excelMergeCell);

excelExportSetting.setHeaderRow(cellList);//设置表头
excelExportSetting.setDataList(datas);//设置数据

//写入excel
ExcelPoiHelp.poiWrite(wb, billInfoSheet, excelExportSetting);

//写入response
String outFile = "outputFile.xls";
response.reset();
response.addHeader("Content-Disposition", "attachment;filename="+ new String(outFile.getBytes()));
OutputStream toClient = new BufferedOutputStream(response.getOutputStream());
response.setContentType("application/vnd.ms-excel;charset=utf-8");
wb.write(toClient);

```

# 目前的功能
* excel导出（支持合并单元格）

# 规划的功能
* excel支持导出公式
* excel生成html

