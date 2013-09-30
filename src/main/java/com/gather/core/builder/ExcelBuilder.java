package com.gather.core.builder;

import com.gather.core.sheet.ISheetCreator;
import com.gather.core.sheet.ISheetPopulator;
import com.gather.core.workbook.IWorkbookCreator;
import org.apache.log4j.Logger;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

/**
 * Created with IntelliJ IDEA.
 * $ Project: excelcommons
 * User: rodrigotroy
 * Date: 9/30/13
 * Time: 12:37 PM
 */
public class ExcelBuilder {
    private static final Logger LOG = Logger.getLogger(ExcelBuilder.class);

    private ISheetCreator sheetCreator;
    private IWorkbookCreator workbookCreator;
    private ISheetPopulator bodySheetPopulator;
    private ISheetPopulator headerSheetPopulator;

    public ExcelBuilder(ISheetCreator sheetCreator,
                        IWorkbookCreator workbookCreator,
                        ISheetPopulator bodySheetPopulator,
                        ISheetPopulator headerSheetPopulator) {
        this.sheetCreator = sheetCreator;
        this.workbookCreator = workbookCreator;
        this.bodySheetPopulator = bodySheetPopulator;
        this.headerSheetPopulator = headerSheetPopulator;
    }

    public void createExcel() {
        if (sheetCreator != null && workbookCreator != null) {
            workbookCreator.createWorkbook();

            sheetCreator.createSheet(workbookCreator.getWorkbook());

            headerSheetPopulator.populate(sheetCreator.getSheet());

            bodySheetPopulator.populate(sheetCreator.getSheet());
        }
    }

    public ByteArrayOutputStream getStream() {
        ByteArrayOutputStream os = new ByteArrayOutputStream();

        try {
            workbookCreator.getWorkbook().write(os);
        } catch (IOException e) {
            LOG.error(e.getMessage());
        } catch (Exception e) {
            LOG.error(e.getMessage());
        }

        return os;
    }
}
