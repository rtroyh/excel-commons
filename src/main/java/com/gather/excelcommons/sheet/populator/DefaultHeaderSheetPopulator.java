package com.gather.excelcommons.sheet.populator;

import com.gather.excelcommons.sheet.populator.columnResolver.DefaultColumnVisibilityResolver;
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
    private IColumnVisibilityResolver columnVisibilityResolver;

    public DefaultHeaderSheetPopulator(IDataTableModel model,
                                       Integer rowStart) {
        this.model = model;
        this.rowStart = rowStart;
        this.columnVisibilityResolver = new DefaultColumnVisibilityResolver();
    }

    public DefaultHeaderSheetPopulator(IDataTableModel model,
                                       Integer rowStart,
                                       IColumnVisibilityResolver columnVisibilityResolver) {
        this.model = model;
        this.rowStart = rowStart;
        this.columnVisibilityResolver = columnVisibilityResolver;
    }

    public DefaultHeaderSheetPopulator(IDataTableModel model,
                                       Integer rowStart,
                                       CellStyle cellStyleHeader,
                                       IColumnVisibilityResolver columnVisibilityResolver) {
        this.model = model;
        this.rowStart = rowStart;
        this.cellStyleHeader = cellStyleHeader;
        this.columnVisibilityResolver = columnVisibilityResolver;
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
    public void populate(Sheet sheet) throws
                                      Exception {
        LOG.info("INICIO POBLAMIENTO SHEET");

        short columnIndex = 0;

        Row headerRow = sheet.createRow(rowStart);

        final CellStyle cellStyle = getCellStyleHeader(sheet.getWorkbook());

        for (List<Object> header : model.getHeaders()) {
            boolean columnaNoesImagen = !header.get(1).equals(5);
            boolean columnaEsVisible = this.columnVisibilityResolver.isVisible(header);

            if (columnaNoesImagen && columnaEsVisible) {
                Cell cell = headerRow.createCell(columnIndex);
                cell.setCellStyle(cellStyle);

                String title = Validator.validateString(header.get(0)) ? header.get(0).toString() : " ";

                cell.setCellValue(title);
                columnIndex++;
            }
        }

        sheet.createFreezePane(0,
                               rowStart + 1);

        for (int x = 0; x < model.getHeaders().size(); x++) {
            sheet.autoSizeColumn(x);
        }
    }
}
