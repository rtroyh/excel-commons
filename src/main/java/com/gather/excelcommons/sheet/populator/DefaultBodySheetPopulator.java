package com.gather.excelcommons.sheet.populator;

import com.gather.excelcommons.sheet.populator.style.NumberCellStyle;
import com.gather.excelcommons.sheet.util.CellUtil;
import com.gather.gathercommons.model.IDataTableModel;
import com.gather.gathercommons.util.Validator;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * Created with IntelliJ IDEA.
 * $ Project: excelcommons
 * User: rodrigotroy
 * Date: 10/28/13
 * Time: 4:18 PM
 */
public class DefaultBodySheetPopulator implements ISheetPopulator {
    private static final Logger LOG = Logger.getLogger(DefaultBodySheetPopulator.class);

    private IDataTableModel model;
    private Integer rowStart;

    private DataFormat dataFormat;
    private CellStyle cellStyleDate;
    private CellStyle cellStylePorcentual;
    private List<NumberCellStyle> numberCellStyles;

    public DefaultBodySheetPopulator(IDataTableModel model,
                                     Integer rowStart) {
        this.model = model;
        this.rowStart = rowStart;
    }

    private DataFormat getDataFormat(final Workbook wb) {
        if (dataFormat == null) {
            dataFormat = wb.createDataFormat();
        }

        return dataFormat;
    }

    private List<NumberCellStyle> getNumberCellStyles() {
        if (numberCellStyles == null) {
            numberCellStyles = Collections.synchronizedList(new ArrayList<NumberCellStyle>());
        }

        return numberCellStyles;
    }

    private CellStyle getNumberCellStyle(final Workbook wb,
                                         final Integer decimals) {
        LOG.debug("INICIO OBTENCION ESTILO PARA CELDA NUMERICA");

        Boolean exists = false;

        for (NumberCellStyle numberCellStyle : this.getNumberCellStyles()) {
            if (numberCellStyle.getDecimalCount().equals(decimals)) {
                exists = true;
            }

            if (exists) {
                return numberCellStyle.getCellStyle();
            }
        }

        if (!exists) {
            CellStyle cellStyle = wb.createCellStyle();

            NumberCellStyle numberCellStyle = new NumberCellStyle(cellStyle,
                                                                  this.getDataFormat(wb),
                                                                  decimals);

            this.getNumberCellStyles().add(numberCellStyle);

            return numberCellStyle.getCellStyle();
        }

        return wb.createCellStyle();
    }

    private CellStyle getCellStylePorcentual(Workbook wb) {
        if (this.cellStylePorcentual == null) {
            this.cellStylePorcentual = wb.createCellStyle();
            this.cellStylePorcentual.setDataFormat(wb.createDataFormat().getFormat("0.00%"));
        }

        return this.cellStylePorcentual;
    }

    private CellStyle getCellStyleDate(Workbook wb) {
        if (this.cellStyleDate == null) {
            this.cellStyleDate = wb.createCellStyle();
            this.cellStyleDate.setDataFormat(wb.createDataFormat().getFormat("YYYY-MM-DD"));
        }

        return this.cellStyleDate;
    }

    @Override
    public void populate(Sheet sheet) {
        LOG.info("INICIO POBLAMIENTO SHEET");

        for (List<Object> row : model.getRows()) {
            Row eRow = sheet.createRow(rowStart);

            int xHeader = 0;
            int xExcel = 0;
            for (List<Object> header : model.getHeaders()) {
                boolean esColumnaVisible = header.get(4).equals(2) || header.get(4).equals(3);
                boolean esTexto = header.get(1).equals(1);
                boolean esNumerico = header.get(1).equals(2);
                boolean esPorcentual = header.get(1).equals(3);
                boolean esFecha = header.get(1).equals(4);
                boolean esImagen = header.get(1).equals(5);

                if (!esImagen && esColumnaVisible) {
                    Object o = row.get(xHeader);

                    if (o != null) {
                        Cell cell = eRow.createCell(xExcel);

                        if (esPorcentual) {
                            LOG.debug("CELDA PORCENTUAL");
                            cell.setCellStyle(getCellStylePorcentual(sheet.getWorkbook()));
                            CellUtil.setNumberValue(o,
                                                    cell);
                        } else if (esFecha) {
                            LOG.debug("CELDA FECHA");
                            CellUtil.setDateValue(o,
                                                  cell,
                                                  getCellStyleDate(sheet.getWorkbook()));
                        } else if (esNumerico) {
                            LOG.debug("CELDA NUMERICA");

                            final int decimals = Validator.validateNumber(header.get(2)) ? (Integer) header.get(2) : 0;

                            cell.setCellStyle(getNumberCellStyle(sheet.getWorkbook(),
                                                                 decimals));
                            CellUtil.setNumberValue(o,
                                                    cell);


                        } else if (esTexto) {
                            LOG.debug("CELDA TEXTO");

                            CellUtil.setStringValue(o,
                                                    cell);
                        }
                    }

                    xExcel++;
                }

                xHeader++;
            }

            rowStart++;
        }

        for (int x = 0; x < model.getHeaders().size(); x++) {
            sheet.autoSizeColumn(x);
        }
    }
}
