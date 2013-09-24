package com.gather.core.header;

import org.apache.poi.ss.usermodel.Row;

/**
 * Created with IntelliJ IDEA.
 * User: rodrigotroy
 * Date: 9/24/13
 * Time: 10:41 AM
 * To change this template use File | Settings | File Templates.
 */
public interface IHeaderBuilder {

    public void createHeader(Row headerRow);
}
