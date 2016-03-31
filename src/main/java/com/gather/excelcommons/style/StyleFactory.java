package com.gather.excelcommons.style;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

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
        cellStyleHeader.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyleHeader.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        cellStyleHeader.setFillForegroundColor(HSSFColor.DARK_BLUE.index);
        cellStyleHeader.setFillPattern(CellStyle.SOLID_FOREGROUND);

        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 11);
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        font.setColor(HSSFColor.WHITE.index);

        cellStyleHeader.setFont(font);

        return cellStyleHeader;
    }
}
