package com.gather.excelcommons.sheet.creator;

import com.gather.gathercommons.util.Validator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;

/**
 * Created by rodrigotroy on 10/27/14.
 */
public class DefaultSheetCreator implements ISheetCreator {
    private XSSFSheet sheet;
    private Object sheetName;
    private Integer zoom;

    public DefaultSheetCreator(Object sheetName,
                               Integer zoom) {
        this.sheetName = sheetName;
        this.zoom = zoom;
    }

    public DefaultSheetCreator(Object sheetName) {
        this.sheetName = sheetName;
        this.zoom = 100;
    }

    @Override
    public XSSFSheet createSheet(XSSFWorkbook workbook) {
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

        this.sheet.setZoom(zoom);

        return sheet;
    }

    @Override
    public final XSSFSheet getSheet() {
        return this.sheet;
    }
}
