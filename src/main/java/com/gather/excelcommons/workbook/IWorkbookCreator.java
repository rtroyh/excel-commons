package com.gather.excelcommons.workbook;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Created with IntelliJ IDEA.
 * $ Project: excelcommons
 * User: rodrigotroy
 * Date: 9/30/13
 * Time: 10:45 AM
 */
public interface IWorkbookCreator {
    public XSSFWorkbook createWorkbook();

    public XSSFWorkbook getWorkbook();
}
