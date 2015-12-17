package com.gather.excelcommons.workbook;

import org.apache.poi.ss.usermodel.Workbook;

/**
 * Created with IntelliJ IDEA.
 * $ Project: excelcommons
 * User: rodrigotroy
 * Date: 9/30/13
 * Time: 10:45 AM
 */
public interface IWorkbookCreator {
    Workbook createWorkbook(String Author);

    Workbook createWorkbook();

    Workbook getWorkbook();
}
