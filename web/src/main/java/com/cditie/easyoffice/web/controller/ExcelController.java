package com.cditie.easyoffice.web.controller;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.cditie.easyoffice.excel.bean.*;
import net.sf.jxls.transformer.XLSTransformer;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import com.cditie.easyoffice.web.App;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.BufferedOutputStream;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

/**
 * Created by zhuyunhui on 9/13/2017.
 */
@RequestMapping("excel")
@Controller
public class ExcelController {
	/**
	 * 日志处理
	 */
	private final Logger logger = LoggerFactory.getLogger(this.getClass());


	@RequestMapping("")
	@ResponseBody
	String excel() {
		return "Hello World!";
	}


	/**
	 * excel 下载
	 */
	@RequestMapping(value = "downLoad")
	public void downLoad(HttpServletRequest request, HttpServletResponse response) throws Exception {

		Workbook wb = null;
		try {
			logger.info(">>>>>>>>ReportViewController.downLoad start>>");

			//=======================================数据
			List<Map<String,Object>> datas = Lists.newArrayList();
			Map<String,Object> data0 = Maps.newHashMap();
			data0.put("date", "2017-01-01");
			data0.put("date1", "2017-01-01");
			Map<String,Object> data1 = Maps.newHashMap();
			data1.put("haodai","100");
			data1.put("daxuedai","100");
			data1.put("zirandai","100");
			data0.put("zidonghebishu",data1);
			datas.add(data0);
			//==========================================


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
							new ExcelCell("好贷","haodai"),
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

		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			wb.close();
		}

	}



}
