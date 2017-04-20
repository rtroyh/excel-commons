package com.gather.excelcommons.sheet.populator.columnResolver;

import com.gather.excelcommons.sheet.populator.IColumnVisibilityResolver;

import java.util.List;

/**
 * Created with IntelliJ IDEA.
 * $ Project: excel-commons
 * User: rodrigotroy
 * Date: 4/20/17
 * Time: 16:33
 */
public class NotZeroColumnVisibilityResolver implements IColumnVisibilityResolver {
    @Override
    public Boolean isVisible(List<Object> headerRow) {
        return !headerRow.get(4).equals(0);
    }
}
