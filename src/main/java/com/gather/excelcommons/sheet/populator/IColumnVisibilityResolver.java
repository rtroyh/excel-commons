package com.gather.excelcommons.sheet.populator;

import java.util.List;

/**
 * Created with IntelliJ IDEA.
 * $ Project: excel-commons
 * User: rodrigotroy
 * Date: 4/20/17
 * Time: 16:28
 */
public interface IColumnVisibilityResolver {
    Boolean isVisible(List<Object> headerRow);
}
