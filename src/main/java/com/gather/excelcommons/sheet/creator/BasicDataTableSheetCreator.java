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
 * Date: 12/13/13
 * Time: 11:08 AM
 */
public class BasicDataTableSheetCreator implements ISheetCreator {
    private IDataTableModel model;
    private Sheet sheet;
    private Object sheetName;

    public BasicDataTableSheetCreator(Object sheetName,
                                      IDataTableModel model) {
        this.sheetName = sheetName;
        this.model = model;
    }

    @Override
    public Sheet createSheet(Workbook workbook) {
        if (Validator.validateDataTableModel(model)) {
            final List<List<Object>> titles = model.getTitles();

            if (Validator.validateList(titles)) {
                final Object valor = sheetName;

                if (Validator.validateString(sheetName)) {
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

        this.sheet.setZoom(1,
                           1);

        return sheet;
    }

    @Override
    public Sheet getSheet() {
        return this.sheet;
    }
}
