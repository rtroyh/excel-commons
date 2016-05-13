package com.gather.excelcommons.sheet.creator;

import com.gather.gathercommons.model.IDataTableModel;
import com.gather.gathercommons.util.Validator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

/**
 * Created with IntelliJ IDEA.
 * $ Project: excelcommons
 * User: rodrigotroy
 * Date: 11/25/13
 * Time: 6:19 PM
 */
public class DataTableSheetCreator implements ISheetCreator {
    private IDataTableModel model;
    private Sheet sheet;

    public DataTableSheetCreator(IDataTableModel model) {
        this.model = model;
    }

    @Override
    public Sheet createSheet(Workbook workbook) {
        if (Validator.validateDataTableModel(model)) {
            final List<List<Object>> titles = model.getTitles();

            if (Validator.validateList(titles)) {
                final Object valor = titles.get(0).get(2);

                if (Validator.validateString(valor)) {
                    String name = valor.toString();
                    name = name.replaceAll("/",
                                           " ");
                    name = name.replaceAll(":",
                                           " ");
                    this.sheet = workbook.createSheet(name);
                }
            }
        }

        if (this.sheet == null) {
            this.sheet = workbook.createSheet();
        }

        this.sheet.setZoom(75);

        return sheet;
    }

    @Override
    public Sheet getSheet() {
        return this.sheet;
    }
}
