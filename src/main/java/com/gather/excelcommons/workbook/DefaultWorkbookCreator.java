package com.gather.excelcommons.workbook;

import org.apache.poi.POIXMLProperties;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Created with IntelliJ IDEA.
 * $ Project: excelcommons
 * User: rodrigotroy
 * Date: 04-10-15
 * Time: 22:36
 */
public class DefaultWorkbookCreator implements IWorkbookCreator {
    private Workbook workbook;

    @Override
    public Workbook createWorkbook(String author) {
        workbook = new XSSFWorkbook();

        POIXMLProperties xmlProps = ((XSSFWorkbook) workbook).getProperties();
        POIXMLProperties.CoreProperties coreProps = xmlProps.getCoreProperties();
        coreProps.setCreator(author);

        return workbook;
    }

    @Override
    public Workbook createWorkbook() {
        workbook = new XSSFWorkbook();

        return workbook;
    }

    @Override
    public Workbook getWorkbook() {
        return workbook;
    }
}
