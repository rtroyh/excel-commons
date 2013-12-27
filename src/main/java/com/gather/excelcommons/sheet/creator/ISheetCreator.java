package com.gather.excelcommons.sheet.creator;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Created with IntelliJ IDEA.
 * User: rodrigotroy
 * Date: 9/24/13
 * Time: 10:45 AM
 * To change this template use File | Settings | File Templates.
 */
public interface ISheetCreator {
    public XSSFSheet createSheet(XSSFWorkbook wb);

    public XSSFSheet getSheet();
}
