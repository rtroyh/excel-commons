package com.gather.excelcommons.sheet.populator;

import org.apache.poi.ss.usermodel.Sheet;

/**
 * Created with IntelliJ IDEA.
 * $ Project: excelcommons
 * User: rodrigotroy
 * Date: 9/30/13
 * Time: 12:19 PM
 */
public interface ISheetPopulator {
    void populate(Sheet sheet) throws
                               Exception;
}
