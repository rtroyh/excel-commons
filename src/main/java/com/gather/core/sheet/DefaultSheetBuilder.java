package com.gather.core.sheet;

import com.gather.gathercommons.model.IDataTableModel;
import com.gather.gathercommons.util.Validator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Created with IntelliJ IDEA.
 * User: rodrigotroy
 * Date: 9/24/13
 * Time: 10:56 AM
 * To change this template use File | Settings | File Templates.
 */
public class DefaultSheetBuilder implements ISheetCreator {
    private IDataTableModel iteracionModel;
    private IDataTableModel model;
    private XSSFSheet sheet;

    public DefaultSheetBuilder(IDataTableModel iteracionModel,
                               IDataTableModel model) {
        this.iteracionModel = iteracionModel;
        this.model = model;
    }

    public XSSFSheet getSheet() {
        return sheet;
    }

    @Override
    public XSSFSheet createSheet(XSSFWorkbook wb) {
        sheet = null;
        if (Validator.validateList(model.getTitles())) {
            if (Validator.validateString(model.getTitles().get(0).get(4))) {
                String name = model.getTitles().get(0).get(4).toString();
                name = name.replaceAll("/",
                                       " ");
                sheet = wb.createSheet(name);
            }

            if (sheet == null) {
                sheet = wb.createSheet();
            }

            boolean existeMensaje = existeMensaje(iteracionModel);

            sheet.createFreezePane(0,
                                   existeMensaje ? 3 : 1,
                                   0,
                                   existeMensaje ? 3 : 1);
        }

        return sheet;
    }

    private boolean existeMensaje(IDataTableModel iteracionModel) {
        boolean existeMensaje = false;

        if (Validator.validateList(iteracionModel.getTitles().get(0),
                                   9)) {
            final Object mensaje = iteracionModel.getTitles().get(0).get(8);
            existeMensaje = Validator.validateString(mensaje);
        }

        return existeMensaje;
    }
}
