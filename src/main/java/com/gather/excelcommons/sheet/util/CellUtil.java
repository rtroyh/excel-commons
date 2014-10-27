package com.gather.excelcommons.sheet.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.sql.Time;

/**
 * Created by rodrigotroy on 10/27/14.
 */
public class CellUtil {
    public static void setStringValue(final Object o,
                                      final Cell cell) {
        if (o instanceof String) {
            cell.setCellValue((String) o);
        } else {
            cell.setCellValue(o.toString());
        }
    }

    public static void setNumberValue(final Object o,
                                      final Cell cell) {
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
            CellUtil.setStringValue(o,
                                    cell);
        }
    }

    public static void setDateValue(final Object o,
                                    final Cell cell,
                                    final CellStyle cellStyle) {
        if (o instanceof Time) {
            cell.setCellValue(o.toString());
        } else if (o instanceof java.sql.Date) {
            cell.setCellStyle(cellStyle);
            cell.setCellValue((java.sql.Date) o);
        } else if (o instanceof java.util.Date) {
            cell.setCellStyle(cellStyle);
            cell.setCellValue((java.util.Date) o);
        } else {
            CellUtil.setStringValue(o,
                                    cell);
        }
    }
}
