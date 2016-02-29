package com.gather.excelcommons.sheet.populator;

import org.apache.poi.ss.usermodel.Sheet;

/**
 * Created with IntelliJ IDEA.
 * $ Project: excelcommons
 * User: rodrigotroy
 * Date: 10/28/13
 * Time: 10:57 AM
 */
public class CompleteSheetPopulator implements ISheetPopulator {
    private ISheetPopulator header;
    private ISheetPopulator body;

    public CompleteSheetPopulator(ISheetPopulator header,
                                  ISheetPopulator body) {
        this.header = header;
        this.body = body;
    }

    @Override
    public void populate(Sheet sheet) throws
                                      Exception {
        header.populate(sheet);
        body.populate(sheet);
    }
}
