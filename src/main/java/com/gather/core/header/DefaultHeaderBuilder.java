package com.gather.core.header;

import com.gather.gathercommons.model.IDataTableModel;
import org.apache.poi.ss.usermodel.Row;

/**
 * Created with IntelliJ IDEA.
 * User: rodrigotroy
 * Date: 9/24/13
 * Time: 11:22 AM
 * To change this template use File | Settings | File Templates.
 */
public class DefaultHeaderBuilder implements IHeaderBuilder {
    private IDataTableModel iteracionModel;
    private IDataTableModel model;

    public DefaultHeaderBuilder(IDataTableModel iteracionModel,
                               IDataTableModel model) {
        this.iteracionModel = iteracionModel;
        this.model = model;
    }
    @Override
    public void createHeader(Row headerRow) {
        //To change body of implemented methods use File | Settings | File Templates.
    }
}
