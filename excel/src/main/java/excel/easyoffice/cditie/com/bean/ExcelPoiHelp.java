package excel.easyoffice.cditie.com.bean;

import java.math.BigDecimal;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.util.StringUtils;


public class ExcelPoiHelp {

    public static int CELLSTYLE_MONEY = 1;
    public static int CELLSTYLE_RATE = 2;
    public static int CELLSTYLE_DATE = 3;
    public static int CELLSTYLE_TEXTBORD = 4;
    public static int CELLSTYLE_TITLE = 0;

    private static CellStyle getCellStyle(int type, Workbook wb, Map<Integer, CellStyle> excelCellStype) {
        if (!excelCellStype.containsKey(type)) {
            excelCellStype = new HashMap<Integer, CellStyle>();
            CellStyle cellStyle = null;
            //1货币，2百分比，3日期，0标题
            cellStyle = wb.getSheet("格式").getRow(type).getCell(1).getCellStyle();
            excelCellStype.put(type, cellStyle);
            return cellStyle;
        } else {
            return excelCellStype.get(type);
        }
    }

    public static boolean isEmpty(Object obj) {
        return obj == null || String.valueOf(obj).trim().length() == 0;
    }


    public static String getText(Object data) {
        if (isEmpty(data))
            return "";
        else
            return String.valueOf(data);
    }


    /**
     * 获取单元格内容
     **/
    public static String getCellValue(Object cellValue) {
        if (cellValue instanceof BigDecimal) {
            BigDecimal bd = (BigDecimal) cellValue;
            return bd.stripTrailingZeros().toPlainString();
        } else {
            return getText(cellValue);
        }

    }

    /**
     * poi导出
     **/
    public static void poiWrite(Workbook wb, Sheet billInfoSheet, ExcelExportSetting setting) {
        Map<Integer, CellStyle> excelCellStype = new HashMap<Integer, CellStyle>();

        if (StringUtils.isEmpty(setting.getDataList())) {
            return;
        }

        //模板移动位置
//		billInfoSheet.shiftRows(setting.getShiftRow(), billInfoSheet.getLastRowNum(), setting.getDataList().size()+1,true,false);

        int headerRowIndex = setting.getBlankRow();
        //title第一行
        Row rowTitle = billInfoSheet.createRow(headerRowIndex);
        List<PoiCell> headerRow = setting.getHeaderRow();

        int titleCell = 0;
        boolean hasMerge = false;
        for (PoiCell poiCell : headerRow) {
            if (poiCell.isMerge()) {
                Cell cell = rowTitle.createCell(titleCell);
                cell.setCellValue(poiCell.getValue());
                cell.setCellStyle(getCellStyle(CELLSTYLE_TITLE, wb, excelCellStype));
                titleCell++;
                ExcelMergeCell mergePoiCell = (ExcelMergeCell) poiCell;
                if (mergePoiCell.getMergeSize() > 1) {
                    for (int i = 0; i < mergePoiCell.getMergeSize() - 1; i++) {
                        cell = rowTitle.createCell(titleCell);
                        cell.setCellValue("");
                        cell.setCellStyle(getCellStyle(CELLSTYLE_TITLE, wb, excelCellStype));
                        titleCell++;
                    }
                }
                hasMerge = true;
            } else {
                Cell cell = rowTitle.createCell(titleCell);
                cell.setCellValue(poiCell.getValue());
                cell.setCellStyle(getCellStyle(CELLSTYLE_TITLE, wb, excelCellStype));
                titleCell++;
            }

        }

        //判断是否有第二行
        if (hasMerge) {
            titleCell = 0;
            headerRowIndex++;
            rowTitle = billInfoSheet.createRow(headerRowIndex);
            for (PoiCell poiCell : headerRow) {
                if (poiCell.isMerge()) {
                    ExcelMergeCell mergePoiCell = (ExcelMergeCell) poiCell;
                    if (mergePoiCell.getMergeSize() > 1) {
                        for (PoiCell excelCell : mergePoiCell.getSubCell()) {
                            Cell cell = rowTitle.createCell(titleCell);
                            cell.setCellValue(excelCell.getValue());
                            cell.setCellStyle(getCellStyle(CELLSTYLE_TITLE, wb, excelCellStype));
                            titleCell++;
                        }

                    } else {
                        Cell cell = rowTitle.createCell(titleCell);
                        cell.setCellValue("");
                        cell.setCellStyle(getCellStyle(CELLSTYLE_TITLE, wb, excelCellStype));
                        titleCell++;

                    }
                }
            }
        }
        //数据项的匹配优先第二行
        List<Map<String, Object>> dataList = setting.getDataList();
        if (!StringUtils.isEmpty(dataList)) {
            int dataCell = 0;
            for (Map<String, Object> data : dataList) {
                dataCell = 0;
                headerRowIndex++;
                Row dataRow = billInfoSheet.createRow(headerRowIndex);
                for (PoiCell poiCell : headerRow) {
                    if (poiCell.isMerge()) {
                        ExcelMergeCell mergePoiCell = (ExcelMergeCell) poiCell;
                        if (mergePoiCell.getMergeSize() >= 1) {
                            Object value = data.get(mergePoiCell.getDatakey());
                            if (value instanceof Map) {
                                @SuppressWarnings("unchecked")
                                Map<String, Object> dataSub = (Map<String, Object>) value;
                                for (PoiCell excelCell : mergePoiCell.getSubCell()) {
                                    Cell cell = dataRow.createCell(dataCell++);
                                    cell.setCellStyle(getCellStyle(CELLSTYLE_TEXTBORD, wb, excelCellStype));
                                    Object subCellValue = dataSub.get(excelCell.getKey());
                                    cell.setCellValue(getCellValue(subCellValue));
                                }
                            }
                        } else {
                            Cell cell = dataRow.createCell(dataCell++);
                            cell.setCellStyle(getCellStyle(CELLSTYLE_TEXTBORD, wb, excelCellStype));
                            Object value = data.get(mergePoiCell.getDatakey());
                            cell.setCellValue(getCellValue(value));
                        }

                    } else {
                        ExcelCell excelCell = (ExcelCell) poiCell;
                        Cell cell = dataRow.createCell(dataCell++);
                        cell.setCellStyle(getCellStyle(CELLSTYLE_TEXTBORD, wb, excelCellStype));
                        Object value = data.get(excelCell.getDatakey());
                        cell.setCellValue(getCellValue(value));
                    }

                }

            }

        }


        //合并单元格处理
        int colomnIdex = 0; // 列起始位置
        int startRowIndex = 0;// 行起始位置
        for (PoiCell poiCell : headerRow) {
            if (poiCell.isMerge()) {
                ExcelMergeCell mergePoiCell = (ExcelMergeCell) poiCell;
                if (mergePoiCell.getMergeSize() > 1) {
                    int subCellSize = mergePoiCell.getSubCell().size();
                    billInfoSheet.addMergedRegion(
                            new CellRangeAddress(startRowIndex, startRowIndex, colomnIdex, colomnIdex + subCellSize - 1));
                    colomnIdex = colomnIdex + subCellSize;

                } else {
                    billInfoSheet.addMergedRegion(
                            new CellRangeAddress(startRowIndex, startRowIndex + 1, colomnIdex, colomnIdex));
                    colomnIdex++;
                }

            }
        }

        //删除格式sheet
        wb.removeSheetAt(0);


    }
}
