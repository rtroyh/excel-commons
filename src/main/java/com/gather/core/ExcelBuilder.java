package com.gather.core;

import com.gather.gathercommons.model.IDataTableModel;
import com.gather.gathercommons.util.Validator;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.List;

public class ExcelBuilder {
    private static final Logger LOG = Logger.getLogger(ExcelBuilder.class);

    public ByteArrayOutputStream getExcelReport(List<IDataTableModel> models) {
        Workbook wb = new HSSFWorkbook();

        for (IDataTableModel model : models) {
            Sheet sheet = createSheet(model,
                                      wb);

            this.buildSheet(model,
                            wb,
                            sheet);
        }

        return buildStream(wb);
    }

    private void buildSheet(IDataTableModel model,
                           Workbook wb,
                           Sheet sheet) {
        buildSheetHeader(model,
                         wb,
                         sheet);

        short y = 1;
        populateSheet(model,
                      wb,
                      sheet,
                      y);

        for (short x = 0; x < model.getHeaders().size() + 2; x++) {
            sheet.autoSizeColumn(x);
        }
    }

    private Sheet createSheet(IDataTableModel model,
                             Workbook wb) {
        Sheet sheet = null;
        if (Validator.validateList(model.getTitles())) {
            if (Validator.validateString(model.getTitles().get(0).get(4))) {
                String name = model.getTitles().get(0).get(4).toString();
                name = name.replaceAll("/",
                                       " ");
                sheet = wb.createSheet(name);
            }
        }

        if (sheet == null) {
            sheet = wb.createSheet();
        }

        sheet.createFreezePane(0,
                               1,
                               0,
                               1);

        return sheet;
    }

    private ByteArrayOutputStream buildStream(Workbook wb) {
        ByteArrayOutputStream os = new ByteArrayOutputStream();

        try {
            wb.write(os);
        } catch (IOException e) {
            LOG.error(e.getMessage());
        } catch (Exception e) {
            LOG.error(e.getMessage());
        }

        return os;
    }

    private void buildSheetHeader(IDataTableModel model,
                                  Workbook wb,
                                  Sheet sheet) {

        short x = 0;
        Row headerRow = sheet.createRow(x);

        for (List<Object> header : model.getHeaders()) {
            if (!header.get(1).equals(5) && header.get(4).equals(1)) {

                Cell cell = this.createCell(wb,
                                            headerRow,
                                            x,
                                            CellStyle.ALIGN_CENTER,
                                            CellStyle.VERTICAL_CENTER);

                final CellStyle cellStyle = cell.getCellStyle();

                Font font = wb.createFont();
                font.setFontHeightInPoints((short) 11);
                font.setBoldweight(Font.BOLDWEIGHT_BOLD);
                font.setColor(HSSFColor.WHITE.index);
                cellStyle.setFillForegroundColor(HSSFColor.DARK_BLUE.index);
                cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

                cellStyle.setFont(font);

                String title = Validator.validateString(header.get(0)) ? header.get(0).toString() : " ";

                cell.setCellValue(title);
                x++;
            }
        }
    }

    private Cell createCell(Workbook wb,
                            Row row,
                            short column,
                            short halign,
                            short valign) {
        Cell cell = row.createCell(column);
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(halign);
        cellStyle.setVerticalAlignment(valign);
        cell.setCellStyle(cellStyle);

        return cell;
    }

    private void populateSheet(IDataTableModel model,
                               Workbook wb,
                               Sheet sheet,
                               short y) {
        short x;
        for (List<Object> row : model.getRows()) {
            Row eRow = sheet.createRow(y);

            int xHeader = 0;
            int xExcel = 0;
            for (List<Object> header : model.getHeaders()) {
                boolean esColumnaVisible = header.get(4).equals(1);
                boolean esTexto = header.get(1).equals(1);
                boolean esNumerico = header.get(1).equals(2);
                boolean esPorcentual = header.get(1).equals(3);
                boolean esImagen = header.get(1).equals(5);

                if (!esImagen && esColumnaVisible) {
                    Object o = row.get(xHeader);

                    if (o != null) {
                        Cell cell = eRow.createCell(xExcel);

                        if (esPorcentual) {
                            cell.setCellValue(Float.valueOf(o.toString()));
                            agregarEstiloPorcentualCelda(wb,
                                                         cell);
                        }
                        if (esNumerico || esPorcentual) {
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

            y++;
        }
    }

    private void agregarEstiloPorcentualCelda(Workbook wb,
                                              Cell cell) {
        CellStyle style = wb.createCellStyle();
        style.setDataFormat(wb.createDataFormat().getFormat("0.00%"));


        cell.setCellStyle(style);
    }
}
