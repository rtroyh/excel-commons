package com.gather.excelcommons;

import com.gather.excelcommons.cell.CellUtil;
import com.gather.excelcommons.sheet.OldDefaultSheetCreator;
import com.gather.excelcommons.sheet.creator.ISheetCreator;
import com.gather.gathercommons.model.IDataTableModel;
import com.gather.gathercommons.util.Validator;
import org.apache.commons.lang3.StringUtils;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.sql.Time;
import java.util.List;

//USADO SOLO POR MODULO "REPORTES" DE GESTION CORREDOR
public class ExcelBuilder {
    private static final Logger LOG = Logger.getLogger(ExcelBuilder.class);

    private CellStyle cellStyleDate;
    private CellStyle cellStyleHeader;

    private CellStyle cellStyleInteger;
    private CellStyle cellStylePorcentual;
    private CellStyle cellStyle1Decimales;
    private CellStyle cellStyle2Decimales;
    private CellStyle cellStyle4Decimales;

    public ByteArrayOutputStream getExcelReport(IDataTableModel iteracionModel,
                                                List<IDataTableModel> models) {
        if (iteracionModel != null && Validator.validateList(models)) {
            Workbook wb = new XSSFWorkbook();

            for (IDataTableModel model : models) {
                ISheetCreator sheetBuilder = new OldDefaultSheetCreator(iteracionModel,
                                                                        model);
                Sheet sheet = sheetBuilder.createSheet(wb);

                if (sheet != null) {
                    this.populateSheet(iteracionModel,
                                       model,
                                       sheet);
                }
            }

            return buildStream(wb);
        }

        return null;
    }

    private void populateSheet(IDataTableModel iteracionModel,
                               IDataTableModel model,
                               Sheet sheet) {
        this.buildSheetHeader(iteracionModel,
                              model,
                              sheet);

        int y = 1;

        boolean existeMensaje = existeMensaje(iteracionModel);

        this.buildSheetBody(model,
                            sheet,
                            existeMensaje ? (y + 2) : y);

        for (short x = 0; x < model.getHeaders().size() + 2; x++) {
            sheet.autoSizeColumn(x);
        }
    }

    private boolean existeMensaje(IDataTableModel iteracionModel) {
        boolean existeMensaje = false;

        if (Validator.validateList(iteracionModel.getTitles()) && Validator.validateList(iteracionModel.getTitles().get(0),
                                                                                         9)) {
            final Object mensaje = getMensaje(iteracionModel);
            existeMensaje = Validator.validateString(mensaje);
        }

        return existeMensaje;
    }

    private String getMensaje(IDataTableModel iteracionModel) {
        return iteracionModel.getTitles().get(0).get(8).toString();
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

    private void buildSheetHeader(IDataTableModel iteracionModel,
                                  IDataTableModel model,
                                  Sheet sheet) {
        int rowIndex = 0;
        short columnIndex = 0;
        boolean existeMensaje = this.existeMensaje(iteracionModel);

        if (existeMensaje) {
            rowIndex = createMensaje(iteracionModel,
                                     sheet,
                                     rowIndex,
                                     columnIndex);
        }

        Row headerRow = sheet.createRow(rowIndex);

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

    private int createMensaje(IDataTableModel iteracionModel,
                              Sheet sheet,
                              int rowIndex,
                              short columnIndex) {
        final String frase = this.getMensaje(iteracionModel);
        final String[] strings = StringUtils.split(frase,
                                                   "|");

        Row headerRow = sheet.createRow(rowIndex);

        Cell cell = this.createCell(sheet.getWorkbook(),
                                    headerRow,
                                    columnIndex,
                                    HorizontalAlignment.LEFT,
                                    VerticalAlignment.TOP);

        final CellStyle cellStyle = cell.getCellStyle();
        Font font = sheet.getWorkbook().createFont();
        font.setFontHeightInPoints((short) 11);
        font.setBold(true);
        font.setColor(IndexedColors.BLACK.getIndex());
        cellStyle.setFont(font);

        StringBuilder stringBuilder = new StringBuilder();
        for (String string : strings) {
            stringBuilder.append(string);
            stringBuilder.append(" ");
        }

        cell.setCellValue(stringBuilder.toString());
        rowIndex++;
        rowIndex++;

        return rowIndex;
    }

    private CellStyle getCellStyleHeader(Workbook wb) {
        if (this.cellStyleHeader == null) {
            this.cellStyleHeader = wb.createCellStyle();
            this.cellStyleHeader.setAlignment(HorizontalAlignment.CENTER);
            this.cellStyleHeader.setVerticalAlignment(VerticalAlignment.CENTER);
            this.cellStyleHeader.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
            this.cellStyleHeader.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            Font font = wb.createFont();
            font.setFontHeightInPoints((short) 11);
            font.setBold(true);
            font.setColor(IndexedColors.WHITE.getIndex());

            this.cellStyleHeader.setFont(font);
        }

        return this.cellStyleHeader;
    }

    private Cell createCell(Workbook wb,
                            Row row,
                            short column,
                            HorizontalAlignment halign,
                            VerticalAlignment valign) {
        Cell cell = row.createCell(column);
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(halign);
        cellStyle.setVerticalAlignment(valign);
        cell.setCellStyle(cellStyle);

        return cell;
    }

    private void buildSheetBody(IDataTableModel model,
                                Sheet sheet,
                                int y) {
        for (List<Object> row : model.getRows()) {
            LOG.debug("Fila n: " + y);
            Row eRow = sheet.createRow(y);

            int xHeader = 0;
            int xExcel = 0;
            for (List<Object> header : model.getHeaders()) {
                boolean esColumnaVisible = header.get(4).equals(1) || header.get(4).equals(3);
                boolean mostrarCeros = header.get(3).equals(1);
                boolean esTexto = header.get(1).equals(1);
                boolean esNumerico = header.get(1).equals(2);
                boolean esPorcentual = header.get(1).equals(3);
                boolean esFecha = header.get(1).equals(4);
                boolean esImagen = header.get(1).equals(5);
                boolean noUsaDecimales = header.get(2).equals(0);

                if (!esImagen && esColumnaVisible) {
                    Object o = row.get(xHeader);

                    if (o != null) {
                        Cell cell = eRow.createCell(xExcel);

                        if (esPorcentual) {
                            cell.setCellStyle(getCellStylePorcentual(sheet.getWorkbook()));
                        }

                        if (esFecha) {
                            if (o instanceof java.sql.Date) {
                                cell.setCellStyle(getCellStyleDate(sheet.getWorkbook()));
                                cell.setCellValue((java.sql.Date) o);
                            } else if (o instanceof Time) {
                                cell.setCellValue(o.toString());
                            } else if (o instanceof java.util.Date) {
                                cell.setCellStyle(getCellStyleDate(sheet.getWorkbook()));
                                cell.setCellValue((java.util.Date) o);
                            }
                        } else if (esNumerico || esPorcentual) {
                            if (!mostrarCeros && Validator.valorEsCero(o)) {
                                cell.setCellValue("");
                            } else {
                                if (noUsaDecimales) {
                                    cell.setCellStyle(this.getCellStyleInteger(sheet.getWorkbook()));
                                } else {
                                    cell.setCellStyle(this.getDecimalStyle(sheet.getWorkbook(),
                                                                           Integer.valueOf(header.get(2).toString())));
                                }

                                CellUtil.setCellValue(cell,
                                                      o);
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

    private CellStyle getCellStyleInteger(Workbook wb) {
        if (this.cellStyleInteger == null) {
            this.cellStyleInteger = wb.createCellStyle();
            this.cellStyleInteger.setDataFormat(wb.createDataFormat().getFormat("#,##"));
        }

        return this.cellStyleInteger;
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

    private CellStyle getCellStyle1Decimales(Workbook wb) {
        if (this.cellStyle1Decimales == null) {
            this.cellStyle1Decimales = wb.createCellStyle();
            this.cellStyle1Decimales.setDataFormat(wb.createDataFormat().getFormat("#,##0.0"));
        }

        return this.cellStyle1Decimales;
    }

    private CellStyle getCellStyle2Decimales(Workbook wb) {
        if (this.cellStyle2Decimales == null) {
            this.cellStyle2Decimales = wb.createCellStyle();
            this.cellStyle2Decimales.setDataFormat(wb.createDataFormat().getFormat("#,##0.00"));
        }

        return this.cellStyle2Decimales;
    }

    private CellStyle getCellStyle4Decimales(Workbook wb) {
        if (this.cellStyle4Decimales == null) {
            this.cellStyle4Decimales = wb.createCellStyle();
            this.cellStyle4Decimales.setDataFormat(wb.createDataFormat().getFormat("#,##0.0000"));
        }

        return this.cellStyle4Decimales;
    }

    private CellStyle getDecimalStyle(Workbook wb,
                                      int decimalCount) {
        if (decimalCount == 1) {
            return this.getCellStyle1Decimales(wb);
        } else if (decimalCount == 4) {
            return this.getCellStyle4Decimales(wb);
        }

        return this.getCellStyle2Decimales(wb);
    }

}
