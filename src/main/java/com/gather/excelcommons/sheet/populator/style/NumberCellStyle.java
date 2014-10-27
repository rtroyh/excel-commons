package com.gather.excelcommons.sheet.populator.style;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;

/**
 * Created by rodrigotroy on 10/27/14.
 */
public class NumberCellStyle {
    private static final Logger LOG = Logger.getLogger(NumberCellStyle.class);

    private CellStyle cellStyle;
    private DataFormat format;
    private Integer decimalCount;

    public NumberCellStyle(CellStyle cellStyle,
                           DataFormat format,
                           Integer decimalCount) {
        this.cellStyle = cellStyle;
        this.format = format;
        this.decimalCount = decimalCount;
        this.buildFormat();
    }

    private void buildFormat() {
        LOG.debug("INICIO CONSTRUCCION FORMATO");
        StringBuilder builder;

        if (decimalCount == 0) {
            builder = new StringBuilder("#,##0");
        } else {
            builder = new StringBuilder("#,##0.0");

            for (int x = 1; x < decimalCount; x++) {
                builder.append("0");
            }
        }

        final short format = this.format.getFormat(builder.toString());
        cellStyle.setDataFormat(format);
    }

    public final Integer getDecimalCount() {
        if (decimalCount == null) {
            decimalCount = 2;
        }

        return decimalCount;
    }

    public final CellStyle getCellStyle() {
        return cellStyle;
    }
}
