package com.gather.excelcommons.sheet.populator;

import com.gather.gathercommons.model.IDataTableModel;
import com.gather.gathercommons.util.Validator;
import org.apache.poi.ss.usermodel.*;

import java.util.List;

/**
 * Created with IntelliJ IDEA.
 * $ Project: excelcommons
 * User: rodrigotroy
 * Date: 10/1/13
 * Time: 10:58 AM
 */
public class BasicSheetPopulator implements ISheetPopulator {
    private IDataTableModel model;
    private CellStyle cellStyleHeader;

    public BasicSheetPopulator(IDataTableModel model) {
        this.model = model;
    }

    @Override
    public void populate(Sheet sheet) {
        short rowIndex = 0;
        short columnIndex = 0;

        final Row headerRow = sheet.createRow(rowIndex);
        final CellStyle cellStyle = getCellStyleHeader(sheet.getWorkbook());

        for (List<Object> header : model.getHeaders()) {
            if (!header.get(1).equals(5) && header.get(4).equals(1)) {
                Cell cell = headerRow.createCell(columnIndex);
                cell.setCellStyle(cellStyle);

                String title = Validator.validateString(header.get(0)) ? header.get(0).toString() : " ";

                cell.setCellValue(title);
                columnIndex++;
            }
        }
    }

    private CellStyle getCellStyleHeader(Workbook wb) {
        if (this.cellStyleHeader == null) {
            this.cellStyleHeader = wb.createCellStyle();
            this.cellStyleHeader.setAlignment(CellStyle.ALIGN_CENTER);
            this.cellStyleHeader.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
            this.cellStyleHeader.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
            this.cellStyleHeader.setFillPattern(CellStyle.SOLID_FOREGROUND);

            Font font = wb.createFont();
            font.setFontHeightInPoints((short) 11);
            font.setBoldweight(Font.BOLDWEIGHT_BOLD);
            font.setColor(IndexedColors.WHITE.getIndex());

            this.cellStyleHeader.setFont(font);
        }

        return this.cellStyleHeader;
    }
}
