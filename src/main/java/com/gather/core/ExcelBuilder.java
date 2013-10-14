package com.gather.core;

import com.gather.core.sheet.DefaultSheetBuilder;
import com.gather.core.sheet.ISheetCreator;
import com.gather.gathercommons.model.IDataTableModel;
import com.gather.gathercommons.util.Validator;
import org.apache.commons.lang3.StringUtils;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.List;

public class ExcelBuilder {
    private static final Logger LOG = Logger.getLogger(ExcelBuilder.class);

    private CellStyle cellStyleHeader;
    private CellStyle cellStylePorcentual;

    public ByteArrayOutputStream getExcelReport(IDataTableModel iteracionModel,
                                                List<IDataTableModel> models) {
        if (iteracionModel != null && Validator.validateList(models)) {
            XSSFWorkbook wb = new XSSFWorkbook();

            for (IDataTableModel model : models) {
                ISheetCreator sheetBuilder = new DefaultSheetBuilder(iteracionModel,
                                                                     model);
                XSSFSheet sheet = sheetBuilder.createSheet(wb);

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

        short y = 1;

        boolean existeMensaje = existeMensaje(iteracionModel);

        this.buildSheetBody(model,
                            sheet,
                            existeMensaje ? (short) (y + 2) : y);

        for (short x = 0; x < model.getHeaders().size() + 2; x++) {
            sheet.autoSizeColumn(x);
        }
    }

    private boolean existeMensaje(IDataTableModel iteracionModel) {
        boolean existeMensaje = false;

        if (Validator.validateList(iteracionModel.getTitles().get(0),
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
        short rowIndex = 0;
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

    private short createMensaje(IDataTableModel iteracionModel,
                                Sheet sheet,
                                short rowIndex,
                                short columnIndex) {
        final String frase = this.getMensaje(iteracionModel);
        final String[] strings = StringUtils.split(frase,
                                                   "|");

        Row headerRow = sheet.createRow(rowIndex);

        Cell cell = this.createCell(sheet.getWorkbook(),
                                    headerRow,
                                    columnIndex,
                                    CellStyle.ALIGN_LEFT,
                                    CellStyle.VERTICAL_TOP);

        final CellStyle cellStyle = cell.getCellStyle();
        Font font = sheet.getWorkbook().createFont();
        font.setFontHeightInPoints((short) 11);
        font.setBoldweight(Font.BOLDWEIGHT_NORMAL);
        font.setColor(HSSFColor.BLACK.index);
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

    private void buildSheetBody(IDataTableModel model,
                                Sheet sheet,
                                short y) {
        CellStyle cellStylePorcentual = getCellStylePorcentual(sheet.getWorkbook());

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
                            cell.setCellStyle(cellStylePorcentual);
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

    private CellStyle getCellStylePorcentual(Workbook wb) {
        if (this.cellStylePorcentual == null) {
            this.cellStylePorcentual = wb.createCellStyle();
            this.cellStylePorcentual.setDataFormat(wb.createDataFormat().getFormat("0.00%"));
        }

        return this.cellStylePorcentual;
    }
}
