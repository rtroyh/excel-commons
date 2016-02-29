package com.gather.excelcommons.sheet.populator;

import com.gather.gathercommons.model.IDataTableModel;
import org.apache.poi.ss.usermodel.*;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.sql.Time;
import java.util.List;

/**
 * Created with IntelliJ IDEA.
 * $ Project: excelcommons
 * User: rodrigotroy
 * Date: 12/13/13
 * Time: 4:26 PM
 */
public class SuperBodySheetPopulator implements ISheetPopulator {
    private IDataTableModel model;
    private Integer rowStart;

    private CellStyle cellStyleDate;
    private CellStyle cellStylePorcentual;

    public SuperBodySheetPopulator(IDataTableModel model,
                                   Integer rowStart) {
        this.model = model;
        this.rowStart = rowStart;
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
        int headerTotalIndex = 0;
        int headerVisibleIndex = 0;
        for (List<Object> header : model.getHeaders()) {
            boolean esColumnaVisible = header.get(4).equals(2) || header.get(4).equals(3);
            boolean esTexto = header.get(1).equals(1);
            boolean esNumerico = header.get(1).equals(2);
            boolean esPorcentual = header.get(1).equals(3);
            boolean esFecha = header.get(1).equals(4);
            boolean esImagen = header.get(1).equals(5);

            CellStyle cellStyle = null;
            if (esPorcentual) {
                cellStyle = getCellStylePorcentual(sheet.getWorkbook());
            } else if (esFecha) {
                cellStyle = getCellStyleDate(sheet.getWorkbook());
            }

            int rowIndex = rowStart;
            if (!esImagen && esColumnaVisible) {
                for (List<Object> row : model.getRows()) {
                    Row eRow;
                    if (headerVisibleIndex == 0) {
                        eRow = sheet.createRow(rowIndex);
                        rowIndex++;
                    } else {
                        eRow = sheet.getRow(rowIndex);
                    }

                    Object o = row.get(headerTotalIndex);

                    if (o != null) {
                        Cell cell = eRow.createCell(headerVisibleIndex);
                        cell.setCellStyle(cellStyle);

                        if (esFecha) {
                            if (o instanceof Time) {
                                cell.setCellValue(o.toString());
                            } else if (o instanceof java.sql.Date) {
                                cell.setCellValue((java.sql.Date) o);
                            } else if (o instanceof java.util.Date) {
                                cell.setCellValue((java.util.Date) o);
                            }
                        } else if (esNumerico || esPorcentual) {
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
                }

                headerVisibleIndex++;
                headerTotalIndex++;
            } else {
                headerTotalIndex++;
            }
        }
    }
}
