package com.gather.excelcommons.sheet.creator;

import com.gather.gathercommons.util.Validator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Created by rodrigotroy on 10/27/14.
 */
public class DefaultSheetCreator implements ISheetCreator {
    private Sheet sheet;
    private Object sheetName;
    private Integer zoomNumerator;
    private Integer zoomDenominator;

    public DefaultSheetCreator(Object sheetName,
                               Integer zoomNumerator,
                               Integer zoomDenominator) {
        this.sheetName = sheetName;
        this.zoomNumerator = zoomNumerator;
        this.zoomDenominator = zoomDenominator;
    }

    public DefaultSheetCreator(Object sheetName) {
        this.sheetName = sheetName;
        this.zoomNumerator = this.zoomDenominator = 1;
    }

    @Override
    public Sheet createSheet(Workbook workbook) {
        if (Validator.validateString(sheetName)) {
            String name = sheetName.toString();
            name = name.replaceAll("/",
                                   " ");
            name = name.replaceAll(":",
                                   " ");
            this.sheet = workbook.createSheet(name);
        }

        if (this.sheet == null) {
            this.sheet = workbook.createSheet();
        }

        this.sheet.setZoom(100);

        return sheet;
    }

    @Override
    public final Sheet getSheet() {
        return this.sheet;
    }
}
