package org.wuwz.poi.core;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.OutputStream;
import java.net.URLEncoder;

/**
 * Created by lifeng on 2018/6/1.
 */
public class POIUtils {
    private static final int mDefaultRowAccessWindowSize = 100;

    public POIUtils() {
    }

    public static SXSSFWorkbook newSXSSFWorkbook(int rowAccessWindowSize) {
        return new SXSSFWorkbook(rowAccessWindowSize);
    }

    public static SXSSFWorkbook newSXSSFWorkbook() {
        return newSXSSFWorkbook(100);
    }

    public static SXSSFSheet newSXSSFSheet(SXSSFWorkbook wb, String sheetName) {
        return (SXSSFSheet)wb.createSheet(sheetName);
    }

    public static SXSSFRow newSXSSFRow(SXSSFSheet sheet, int index) {
        return (SXSSFRow)sheet.createRow(index);
    }

    public static SXSSFCell newSXSSFCell(SXSSFRow row, int index) {
        return (SXSSFCell)row.createCell(index);
    }

    public static void setColumnWidth(SXSSFSheet sheet, int index, short width, String value) {
        if(width == -1 && value != null && !"".equals(value)) {
            sheet.setDefaultRowHeight((short) (2 * 256));
            sheet.setColumnWidth(index, (short)(value.length() * 512));
        } else {
            width = width == -1?200:width;
            sheet.setColumnWidth(index, (short)((int)((double)width * 35.7D)));
        }

    }

    public static void writeByLocalOrBrowser(HttpServletResponse response, String fileName, SXSSFWorkbook wb, OutputStream out) throws Exception {
        if(response != null) {
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setHeader("Content-disposition", "attachment; filename=" + URLEncoder.encode(String.format("%s%s", new Object[]{fileName, ".xlsx"}), "UTF-8"));
            if(out == null) {
                out = response.getOutputStream();
            }
        }

        wb.write((OutputStream)out);
        ((OutputStream)out).flush();
        ((OutputStream)out).close();
    }

    public static SXSSFSheet setHSSFValidation(SXSSFSheet sheet, String[] textlist, int firstRow, int endRow, int firstCol, int endCol) {
        DataValidationHelper validationHelper = sheet.getDataValidationHelper();
        DataValidationConstraint explicitListConstraint = validationHelper.createExplicitListConstraint(textlist);
        CellRangeAddressList regions = new CellRangeAddressList(firstRow, endRow, firstCol, endCol);
        DataValidation validation = validationHelper.createValidation(explicitListConstraint, regions);
        validation.setSuppressDropDownArrow(true);
        validation.createErrorBox("tip", "请从下拉列表选取");
        validation.setShowErrorBox(true);
        sheet.addValidationData(validation);
        return sheet;
    }

    public static void checkExcelFile(File file) {
        if(file != null && file.exists()) {
            checkExcelFile(file.getAbsolutePath());
        } else {
            throw new IllegalArgumentException("excel文件不存在.");
        }
    }

    public static void checkExcelFile(String file) {
        if(!file.endsWith(".xlsx")) {
            throw new IllegalArgumentException("抱歉,目前ExcelKit仅支持.xlsx格式的文件.");
        }
    }
}
