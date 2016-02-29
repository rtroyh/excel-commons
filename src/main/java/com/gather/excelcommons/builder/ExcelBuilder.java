package com.gather.excelcommons.builder;

import com.gather.excelcommons.sheet.ISheetBuilder;
import com.gather.excelcommons.workbook.IWorkbookCreator;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.ByteArrayOutputStream;
import java.util.List;

/**
 * Created with IntelliJ IDEA.
 * $ Project: excelcommons
 * User: rodrigotroy
 * Date: 9/30/13
 * Time: 12:37 PM
 */
public class ExcelBuilder {
    private static final Logger LOG = Logger.getLogger(ExcelBuilder.class);

    private IWorkbookCreator workbookCreator;
    private List<ISheetBuilder> sheetBuilders;

    public ExcelBuilder(IWorkbookCreator workbookCreator,
                        List<ISheetBuilder> sheetBuilders) {
        this.workbookCreator = workbookCreator;
        this.sheetBuilders = sheetBuilders;
    }

    public void createExcel() throws
                              Exception {
        LOG.info("INICIO CONSTRUCCION ARCHIVO EXCEL");

        if (workbookCreator != null) {
            for (ISheetBuilder sheetBuilder : sheetBuilders) {
                Sheet sheet = sheetBuilder.createSheet(workbookCreator.getWorkbook());
                sheetBuilder.populate(sheet);
            }
        }

        LOG.info("FIN CONSTRUCCION ARCHIVO EXCEL");
    }

    public ByteArrayOutputStream getStream() throws
                                             Exception {
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        workbookCreator.getWorkbook().write(os);

        return os;
    }
}
