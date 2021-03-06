package com.gather.excelcommons.sheet;

import com.gather.excelcommons.sheet.creator.ISheetCreator;
import com.gather.gathercommons.model.IDataTableModel;
import com.gather.gathercommons.util.Validator;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Created with IntelliJ IDEA.
 * User: rodrigotroy
 * Date: 9/24/13
 * Time: 10:56 AM
 * To change this template use File | Settings | File Templates.
 */
public class OldDefaultSheetCreator implements ISheetCreator {
    private static final Logger LOG = Logger.getLogger(OldDefaultSheetCreator.class);

    private IDataTableModel iteracionModel;
    private IDataTableModel model;
    private Sheet sheet;

    public OldDefaultSheetCreator(IDataTableModel iteracionModel,
                                  IDataTableModel model) {
        this.iteracionModel = iteracionModel;
        this.model = model;
    }

    public Sheet getSheet() {
        return sheet;
    }

    @Override
    public Sheet createSheet(Workbook wb) {
        sheet = null;
        if (Validator.validateList(model.getTitles())) {
            if (Validator.validateString(model.getTitles().get(0).get(4))) {
                String name = model.getTitles().get(0).get(4).toString();
                name = name.replaceAll("/",
                                       " ");
                LOG.info("Nombre sheet: " + name);
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


        if (Validator.validateList(iteracionModel.getTitles()) && Validator.validateList(iteracionModel.getTitles().get(0),
                                                                                         9)) {
            final Object mensaje = iteracionModel.getTitles().get(0).get(8);
            existeMensaje = Validator.validateString(mensaje);
        }

        return existeMensaje;
    }
}
