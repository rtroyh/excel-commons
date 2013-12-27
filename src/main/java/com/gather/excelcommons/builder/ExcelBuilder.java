package com.gather.excelcommons.builder;

import com.gather.excelcommons.sheet.ISheetBuilder;
import com.gather.excelcommons.workbook.IWorkbookCreator;
import org.apache.log4j.Logger;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
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

    public void createExcel() {
        LOG.info("INICIO CONSTRUCCION ARCHIVO EXCEL");

        if (workbookCreator != null) {
            workbookCreator.createWorkbook();

            for (ISheetBuilder sheetBuilder : sheetBuilders) {
                XSSFSheet xssfSheet = sheetBuilder.createSheet(workbookCreator.getWorkbook());
                sheetBuilder.populate(xssfSheet);
            }
        }

        LOG.info("FIN CONSTRUCCION ARCHIVO EXCEL");
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
