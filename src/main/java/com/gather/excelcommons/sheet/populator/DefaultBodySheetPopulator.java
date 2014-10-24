package com.gather.excelcommons.sheet.populator;

import com.gather.gathercommons.model.IDataTableModel;
import com.gather.gathercommons.util.Validator;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.sql.Time;
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
    private IDataTableModel model;
    private Integer rowStart;

    private CellStyle cellStyleDate;
    private CellStyle cellStylePorcentual;
    private List<CellStyle> cellStyles;

    public DefaultBodySheetPopulator(IDataTableModel model,
                                     Integer rowStart) {
        this.model = model;
        this.rowStart = rowStart;
    }

    private List<CellStyle> getCellStyles() {
        if (cellStyles == null) {
            cellStyles = Collections.synchronizedList(new ArrayList<CellStyle>());
        }

        return cellStyles;
    }

    private CellStyle getNumberCellStyle(final Workbook wb,
                                         final Integer decimals) {
        CellStyle cellStyle = wb.createCellStyle();

        for (CellStyle style : this.getCellStyles()) {

        }

        return cellStyle;
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
    public void populate(XSSFSheet sheet) {
        for (List<Object> row : model.getRows()) {
            Row eRow = sheet.createRow(rowStart);

            int xHeader = 0;
            int xExcel = 0;
            for (List<Object> header : model.getHeaders()) {
                boolean esColumnaVisible = header.get(4).equals(1) || header.get(4).equals(3);
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
                            cell.setCellStyle(getCellStylePorcentual(sheet.getWorkbook()));
                        }

                        if (esFecha) {
                            if (o instanceof Time) {
                                cell.setCellValue(o.toString());
                            } else if (o instanceof java.sql.Date) {
                                cell.setCellStyle(getCellStyleDate(sheet.getWorkbook()));
                                cell.setCellValue((java.sql.Date) o);
                            } else if (o instanceof java.util.Date) {
                                cell.setCellStyle(getCellStyleDate(sheet.getWorkbook()));
                                cell.setCellValue((java.util.Date) o);
                            } else if (o instanceof Number) {

                            }
                        } else if (esNumerico || esPorcentual) {
                            final int numeroDecimales = Validator.validateNumber(header.get(2)) ? (Integer) header.get(2) : 0;

                            if (o instanceof Double) {
                                cell.setCellValue((Double) o);
                            } else if (o instanceof Integer) {
                                cell.setCellValue((Integer) o);
                            } else if (o instanceof BigDecimal) {
                                cell.setCellValue(((BigDecimal) o).doubleValue());
                            } else if (o instanceof BigInteger) {
                                cell.setCellValue(((BigInteger) o).longValue());
                            } else if (o instanceof Boolean) {
                                cell.setCellValue((Boolean) o);
                            } else if (o instanceof Short) {
                                cell.setCellValue((Short) o);
                            } else if (o instanceof Float) {
                                cell.setCellValue((Float) o);
                            } else if (o instanceof Long) {
                                cell.setCellValue((Long) o);
                            } else {
                                cell.setCellValue(o.toString());
                            }
                        } else if (esTexto) {
                            if (o instanceof String) {
                                cell.setCellValue((String) o);
                            } else {
                                cell.setCellValue(o.toString());
                            }
                        }
                    }

                    xExcel++;
                }

                xHeader++;
            }

            rowStart++;
        }
    }
}
