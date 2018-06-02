package org.wuwz.poi.hanlder;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.wuwz.poi.pojo.ExportItem;

/**
 * Created by lifeng on 2018/6/1.
 */
public interface ExportHandler {
    CellStyle headCellStyle(SXSSFWorkbook var1, ExportItem exportItem);

    String exportFileName(String var1);
}
