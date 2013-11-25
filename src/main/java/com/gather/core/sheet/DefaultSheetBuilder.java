package com.gather.core.sheet;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Created with IntelliJ IDEA.
 * $ Project: excelcommons
 * User: rodrigotroy
 * Date: 10/28/13
 * Time: 10:50 AM
 */
public class DefaultSheetBuilder implements ISheetBuilder {
    private ISheetCreator sheetCreator;
    private ISheetPopulator sheetPopulator;

    public DefaultSheetBuilder(ISheetCreator sheetCreator,
                               ISheetPopulator sheetPopulator) {
        this.sheetCreator = sheetCreator;
        this.sheetPopulator = sheetPopulator;
    }

    @Override
    public XSSFSheet createSheet(XSSFWorkbook wb) {
        return sheetCreator.createSheet(wb);
    }

    @Override
    public XSSFSheet getSheet() {
        return sheetCreator.getSheet();
    }

    @Override
    public void populate(XSSFSheet sheet) {
        sheetPopulator.populate(sheet);
    }
}
