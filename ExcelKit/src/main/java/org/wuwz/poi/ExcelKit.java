package org.wuwz.poi;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import javax.servlet.http.HttpServletResponse;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.wuwz.poi.annotation.ExportConfig;
import org.wuwz.poi.convert.ExportConvert;
import org.wuwz.poi.convert.ExportRange;
import org.wuwz.poi.core.POIUtils;
import org.wuwz.poi.core.XlsxReader;
import org.wuwz.poi.hanlder.ExportHandler;
import org.wuwz.poi.hanlder.ReadHandler;
import org.wuwz.poi.pojo.ExportItem;

public class ExcelKit
{
    private static Logger log = Logger.getLogger(ExcelKit.class);
    private Class<?> mClass = null;
    private HttpServletResponse mResponse = null;
    private String mEmptyCellValue = null;
    private Integer mMaxSheetRecords = Integer.valueOf(10000);
    private Map<String, ExportConvert> mConvertInstanceCache = new HashMap();

    protected ExcelKit() {}

    protected ExcelKit(Class<?> clazz)
    {
        this(clazz, null);
    }

    protected ExcelKit(Class<?> clazz, HttpServletResponse response)
    {
        this.mResponse = response;
        this.mClass = clazz;
    }

    public static ExcelKit $Builder(Class<?> clazz)
    {
        return new ExcelKit(clazz);
    }

    public static ExcelKit $Export(Class<?> clazz, HttpServletResponse response)
    {
        return new ExcelKit(clazz, response);
    }

    public static ExcelKit $Import()
    {
        return new ExcelKit();
    }

    public ExcelKit setEmptyCellValue(String emptyCellValue)
    {
        this.mEmptyCellValue = emptyCellValue;
        return this;
    }

    public ExcelKit setMaxSheetRecords(Integer size)
    {
        this.mMaxSheetRecords = size;
        return this;
    }

    public boolean toExcel(List<?> data, String sheetName)
    {
        required$ExportParams();
        try
        {
            return toExcel(data, sheetName, this.mResponse.getOutputStream());
        }
        catch (IOException e)
        {
            log.error("导出Excel失败:" + e.getMessage(), e);
        }
        return false;
    }

    public boolean toExcel(List<?> data, String sheetName, OutputStream out)
    {
        return toExcel(data, sheetName, new ExportHandler()
        {
            public CellStyle headCellStyle(SXSSFWorkbook wb, ExportItem exportItem)
            {
                CellStyle cellStyle = wb.createCellStyle();
                Font font = wb.createFont();
//                cellStyle.setFillForegroundColor((short)12);
                cellStyle.setFillPattern((short)1);
                cellStyle.setBorderTop((short)1);
                cellStyle.setBorderRight((short)1);
                cellStyle.setBorderBottom((short)1);
                cellStyle.setBorderLeft((short)1);
                cellStyle.setAlignment((short)1);
                cellStyle.setFillForegroundColor((short)17);
                cellStyle.setFillBackgroundColor((short)17);
                font.setBoldweight((short)400);
                font.setColor((short)9);
                cellStyle.setFont(font);
                DataFormat dataFormat = wb.createDataFormat();
                cellStyle.setDataFormat(dataFormat.getFormat("@"));
                return cellStyle;
            }

            public String exportFileName(String sheetName)
            {
                return String.format("导出-%s-%s", new Object[] { sheetName, Long.valueOf(System.currentTimeMillis()) });
            }
        }, out);
    }

    public boolean toExcel(List<?> data, String sheetName, ExportHandler handler, OutputStream out)
    {
        required$BuilderParams();
        long begin = System.currentTimeMillis();
        if ((data == null) || (data.size() < 1))
        {
            log.error("没有检测到数据,不执行导出操作。");
            return false;
        }
        log.info(String.format("即将导出excel数据：%s条,请稍后..", new Object[] { Integer.valueOf(data.size()) }));


        ExportConfig currentExportConfig = null;
        ExportItem currentExportItem = null;
        List<ExportItem> exportItems = new ArrayList();
        for (Field field : this.mClass.getDeclaredFields())
        {
            currentExportConfig = (ExportConfig)field.getAnnotation(ExportConfig.class);
            if (currentExportConfig != null)
            {
                currentExportItem = new ExportItem()
                        .setField(field.getName())
                        .setDisplay("field".equals(currentExportConfig.value()) ? field.getName() : currentExportConfig.value())
                        .setWidth(currentExportConfig.width())
                        .setConvert(currentExportConfig.convert())
                        .setColor(currentExportConfig.color())
                        .setRange(currentExportConfig.range())
                        .setReplace(currentExportConfig.replace())
                        .setDataType(currentExportConfig.dataType());
                exportItems.add(currentExportItem);
            }
            currentExportItem = null;
            currentExportConfig = null;
        }
        SXSSFWorkbook wb = POIUtils.newSXSSFWorkbook();

        double sheetNo = Math.ceil(data.size() / this.mMaxSheetRecords.intValue());
        for (int index = 0; index <= (sheetNo == 0.0D ? sheetNo : sheetNo - 1.0D); index++)
        {
            SXSSFSheet sheet = POIUtils.newSXSSFSheet(wb, sheetName + (index == 0 ? "" : new StringBuilder().append("_").append(index).toString()));
            SXSSFRow headerRow = POIUtils.newSXSSFRow(sheet, 0);
            for (int i = 0; i < exportItems.size(); i++)
            {
                SXSSFCell cell = POIUtils.newSXSSFCell(headerRow, i);
                POIUtils.setColumnWidth(sheet, i, ((ExportItem)exportItems.get(i)).getWidth(), ((ExportItem)exportItems.get(i)).getDisplay());
                cell.setCellValue(((ExportItem)exportItems.get(i)).getDisplay());
                CellStyle style = handler.headCellStyle(wb,exportItems.get(i));
                if (style != null) {
                    cell.setCellStyle(style);
                }
                String range = ((ExportItem)exportItems.get(i)).getRange();
                if (!"".equals(range))
                {
                    String[] ranges = rangeCellValues(range);
                    POIUtils.setHSSFValidation(sheet, ranges, 1, data.size(), i, i);
                }
            }
            if (data.size() > 0)
            {
                int startNo = index * this.mMaxSheetRecords.intValue();
                int endNo = Math.min(startNo + this.mMaxSheetRecords.intValue(), data.size());
                for (int i = startNo; i < endNo; i++)
                {
                    SXSSFRow bodyRow = POIUtils.newSXSSFRow(sheet, i + 1 - startNo);
                    for (int j = 0; j < exportItems.size(); j++)
                    {
                        String cellValue = ((ExportItem)exportItems.get(j)).getReplace();
                        if ("".equals(cellValue)) {
                            try
                            {
                                cellValue = BeanUtils.getProperty(data.get(i), ((ExportItem)exportItems.get(j)).getField());
                            }
                            catch (Exception e)
                            {
                                log.error("获取" + ((ExportItem)exportItems.get(j)).getField() + "的值失败.", e);
                            }
                        }
                        if (!"".equals(((ExportItem)exportItems.get(j)).getConvert())) {
                            cellValue = convertCellValue(Integer.valueOf(Integer.parseInt(cellValue)), ((ExportItem)exportItems.get(j)).getConvert());
                        }
                        POIUtils.setColumnWidth(sheet, j, ((ExportItem)exportItems.get(j)).getWidth(), cellValue);
                        SXSSFCell cell = POIUtils.newSXSSFCell(bodyRow, j);
                        if(!"".equals(exportItems.get(j).getDataType())){
                            if(null != cellValue){
                                try{
                                    BigDecimal a = new BigDecimal(cellValue);
                                    cell.setCellValue(a.doubleValue());
                                }catch (Exception e){
                                    log.error(e.toString() + ":" + cellValue);
                                    cell.setCellValue("");
                                }
                            }else{
                                cell.setCellValue("");
                            }
                        }else {
                            cell.setCellValue("".equals(cellValue) ? null : cellValue);
                        }
                        CellStyle style = wb.createCellStyle();
                        Font font = wb.createFont();
                        font.setColor(((ExportItem)exportItems.get(j)).getColor());
                        style.setFont(font);
                        DataFormat dataFormat = wb.createDataFormat();
                        if(!"".equals(exportItems.get(j).getDataType())){
                            //这种货币类型只能支持值为double
                            style.setDataFormat(dataFormat.getFormat("#,##0.00"));
                        }else {
                            style.setDataFormat(dataFormat.getFormat("@"));
                        }
                        cell.setCellStyle(style);
                    }
                }
            }
        }
        try
        {
            POIUtils.writeByLocalOrBrowser(this.mResponse, handler.exportFileName(sheetName), wb, out);
        }
        catch (Exception e)
        {
            log.error("生成Excel文件失败:" + e.getMessage(), e);
            return false;
        }
        log.info(String.format("Excel处理完成,共生成数据:%s行 (不包含表头),耗时：%s seconds.", new Object[] { Integer.valueOf(data != null ? data.size() : 0),
                Float.valueOf((float)(System.currentTimeMillis() - begin) / 1000.0F) }));
        return true;
    }

    public void readExcel(File excelFile, ReadHandler handler)
    {
        readExcel(excelFile, -1, handler);
    }

    public void readExcel(InputStream is, String fileName, ReadHandler handler)
    {
        readExcel(is, fileName, -1, handler);
    }

    public void readExcel(InputStream is, String fileName, int sheetIndex, ReadHandler handler)
    {
        long begin = System.currentTimeMillis();
        XlsxReader reader = new XlsxReader(handler).setEmptyCellValue(this.mEmptyCellValue);
        try
        {
            if (sheetIndex >= 0) {
                reader.process(is, fileName, sheetIndex);
            } else {
                reader.process(is, fileName);
            }
        }
        catch (Exception e)
        {
            log.error("读取excel文件时发生异常：" + e.getMessage(), e);
        }
        log.info(String.format("Excel读取并处理完成,耗时：%s seconds.", new Object[] { Float.valueOf((float)(System.currentTimeMillis() - begin) / 1000.0F) }));
    }

    public void readExcel(File excelFile, int sheetIndex, ReadHandler handler)
    {
        long begin = System.currentTimeMillis();
        String fileName = excelFile.getAbsolutePath();
        XlsxReader reader = new XlsxReader(handler).setEmptyCellValue(this.mEmptyCellValue);
        try
        {
            if (sheetIndex >= 0) {
                reader.process(fileName, sheetIndex);
            } else {
                reader.process(fileName);
            }
        }
        catch (Exception e)
        {
            log.error("读取excel文件时发生异常：" + e.getMessage(), e);
        }
        log.info(String.format("Excel读取并处理完成,耗时：%s seconds.", new Object[] { Float.valueOf((float)(System.currentTimeMillis() - begin) / 1000.0F) }));
    }

    private String convertCellValue(Integer oldValue, String format)
    {
        try
        {
            String protocol = format.split(":")[0];
            if ("s".equalsIgnoreCase(protocol))
            {
                String[] pattern = format.split(":")[1].split(",");
                for (String p : pattern)
                {
                    String[] cp = p.split("=");
                    if (Integer.parseInt(cp[0]) == oldValue.intValue()) {
                        return cp[1];
                    }
                }
            }
            if ("c".equalsIgnoreCase(protocol))
            {
                String clazz = format.split(":")[1];
                ExportConvert export = (ExportConvert)this.mConvertInstanceCache.get(clazz);
                if (export == null)
                {
                    export = (ExportConvert)Class.forName(clazz).newInstance();
                    this.mConvertInstanceCache.put(clazz, export);
                }
                if (this.mConvertInstanceCache.size() > 10) {
                    this.mConvertInstanceCache.clear();
                }
                return export.handler(oldValue);
            }
        }
        catch (Exception e)
        {
            log.error("出现问题,可能是@ExportConfig.format()的值不规范导致。", e);
        }
        return String.valueOf(oldValue);
    }

    private String[] rangeCellValues(String format)
    {
        try
        {
            String protocol = format.split(":")[0];
            if ("c".equalsIgnoreCase(protocol))
            {
                String clazz = format.split(":")[1];
                ExportRange export = (ExportRange)Class.forName(clazz).newInstance();
                if (export != null) {
                    return export.handler();
                }
            }
        }
        catch (Exception e)
        {
            log.error("出现问题,可能是@ExportConfig.range()的值不规范导致。", e);
        }
        return new String[0];
    }

    private void required$BuilderParams()
    {
        if (this.mClass == null) {
            throw new IllegalArgumentException("请先使用org.wuwz.poi.ExcelKit.$Builder(Class<?>)构造器初始化参数。");
        }
    }

    private void required$ExportParams()
    {
        if ((this.mClass == null) || (this.mResponse == null)) {
            throw new IllegalArgumentException("请先使用org.wuwz.poi.ExcelKit.$Export(Class<?>, HttpServletResponse)构造器初始化参数。");
        }
    }
}
