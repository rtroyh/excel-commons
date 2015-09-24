package com.gather.excelcommons.sheet.creator;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Created with IntelliJ IDEA.
 * User: rodrigotroy
 * Date: 9/24/13
 * Time: 10:45 AM
 * To change this template use File | Settings | File Templates.
 */
public interface ISheetCreator {
    public Sheet createSheet(Workbook wb);

    public Sheet getSheet();
}
