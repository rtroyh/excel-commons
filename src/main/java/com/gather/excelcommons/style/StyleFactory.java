package com.gather.excelcommons.style;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;

/**
 * Created with IntelliJ IDEA.
 * $ Project: excel-commons
 * User: rodrigotroy
 * Date: 31-03-16
 * Time: 13:44
 */
public class StyleFactory {
    public static CellStyle getBlueHeaderCellStyle(Workbook wb) {
        CellStyle cellStyleHeader = wb.createCellStyle();
        cellStyleHeader.setAlignment(HorizontalAlignment.CENTER);
        cellStyleHeader.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyleHeader.setFillForegroundColor(HSSFColor.HSSFColorPredefined.DARK_BLUE.getIndex());
        cellStyleHeader.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 11);
        font.setBold(true);
        font.setColor(HSSFColor.HSSFColorPredefined.WHITE.getIndex());

        cellStyleHeader.setFont(font);

        return cellStyleHeader;
    }
}
