package com.gather.core.sheet;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Created with IntelliJ IDEA.
 * User: rodrigotroy
 * Date: 9/24/13
 * Time: 10:45 AM
 * To change this template use File | Settings | File Templates.
 */
public interface ISheetBuilder {
    public Sheet createSheet(Workbook wb);
}
