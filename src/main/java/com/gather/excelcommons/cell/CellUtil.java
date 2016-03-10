package com.gather.excelcommons.cell;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.sql.Time;

/**
 * Created with IntelliJ IDEA.
 * $ Project: excel-commons
 * User: rodrigotroy
 * Date: 10-03-16
 * Time: 14:55
 */
public class CellUtil {
    public static void setCellValue(CellStyle dateCellStyle,
                                    Cell cell,
                                    Object o) {
        if (o instanceof Time) {
            cell.setCellValue(o.toString());
        } else if (o instanceof java.sql.Date) {
            cell.setCellStyle(dateCellStyle);
            cell.setCellValue((java.sql.Date) o);
        } else if (o instanceof java.util.Date) {
            cell.setCellStyle(dateCellStyle);
            cell.setCellValue((java.util.Date) o);
        }
    }

    public static void setCellValue(Cell cell,
                                    Object o) {
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
    }
}
