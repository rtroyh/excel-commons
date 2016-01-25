package com.gather.excelcommons.sheet.populator;

import com.gather.gathercommons.model.IDataTableModel;
import com.gather.gathercommons.util.Validator;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;

import java.util.List;

/**
 * Created with IntelliJ IDEA.
 * $ Project: excelcommons
 * User: rodrigotroy
 * Date: 11/25/13
 * Time: 6:26 PM
 */
public class DefaultHeaderSheetPopulator implements ISheetPopulator {
    private static final Logger LOG = Logger.getLogger(DefaultHeaderSheetPopulator.class);

    private IDataTableModel model;
    private Integer rowStart;
    private CellStyle cellStyleHeader;

    public DefaultHeaderSheetPopulator(IDataTableModel model,
                                       Integer rowStart) {
        this.model = model;
        this.rowStart = rowStart;
    }

    private CellStyle getCellStyleHeader(Workbook wb) {
        if (this.cellStyleHeader == null) {
            this.cellStyleHeader = wb.createCellStyle();
            this.cellStyleHeader.setAlignment(CellStyle.ALIGN_CENTER);
            this.cellStyleHeader.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
            this.cellStyleHeader.setFillForegroundColor(HSSFColor.DARK_BLUE.index);
            this.cellStyleHeader.setFillPattern(CellStyle.SOLID_FOREGROUND);

            Font font = wb.createFont();
            font.setFontHeightInPoints((short) 11);
            font.setBoldweight(Font.BOLDWEIGHT_BOLD);
            font.setColor(HSSFColor.WHITE.index);

            this.cellStyleHeader.setFont(font);
        }

        return this.cellStyleHeader;
    }

    @Override
    public void populate(Sheet sheet) {
        LOG.info("INICIO POBLAMIENTO SHEET");

        short columnIndex = 0;

        Row headerRow = sheet.createRow(rowStart);

        final CellStyle cellStyle = getCellStyleHeader(sheet.getWorkbook());

        for (List<Object> header : model.getHeaders()) {
            if (!header.get(1).equals(5) && (header.get(4).equals(2) || header.get(4).equals(3))) {
                Cell cell = headerRow.createCell(columnIndex);
                cell.setCellStyle(cellStyle);

                String title = Validator.validateString(header.get(0)) ? header.get(0).toString() : " ";

                cell.setCellValue(title);
                columnIndex++;
            }
        }

        sheet.createFreezePane(0,
                               1);

        for (int x = 0; x < model.getHeaders().size(); x++) {
            sheet.autoSizeColumn(x);
        }
    }
}
