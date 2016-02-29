package com.gather.excelcommons.sheet;

import com.gather.excelcommons.sheet.creator.ISheetCreator;
import com.gather.excelcommons.sheet.populator.ISheetPopulator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Created with IntelliJ IDEA.
 * $ Project: excelcommons
 * User: rodrigotroy
 * Date: 10/28/13
 * Time: 10:50 AM
 */
public class SheetBuilder implements ISheetBuilder {
    private ISheetCreator sheetCreator;
    private ISheetPopulator sheetPopulator;

    public SheetBuilder(ISheetCreator sheetCreator,
                        ISheetPopulator sheetPopulator) {
        this.sheetCreator = sheetCreator;
        this.sheetPopulator = sheetPopulator;
    }

    @Override
    public Sheet createSheet(Workbook wb) {
        return sheetCreator.createSheet(wb);
    }

    @Override
    public Sheet getSheet() {
        return sheetCreator.getSheet();
    }

    @Override
    public void populate(Sheet sheet) throws
                                      Exception {
        sheetPopulator.populate(sheet);
    }
}
