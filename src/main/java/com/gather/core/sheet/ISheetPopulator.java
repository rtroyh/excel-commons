package com.gather.core.sheet;

import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 * Created with IntelliJ IDEA.
 * $ Project: excelcommons
 * User: rodrigotroy
 * Date: 9/30/13
 * Time: 12:19 PM
 */
public interface ISheetPopulator {
    public void populate(XSSFSheet sheet);
}
