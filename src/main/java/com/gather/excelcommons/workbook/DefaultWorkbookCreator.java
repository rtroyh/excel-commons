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
    private String author;

    public DefaultWorkbookCreator() {
    }

    public DefaultWorkbookCreator(String author) {
        this.author = author;
    }

    @Override
    public Workbook getWorkbook() {
        if (workbook == null) {
            if (this.author == null) {
                workbook = new XSSFWorkbook();
            } else {
                workbook = new XSSFWorkbook();

                POIXMLProperties xmlProps = ((XSSFWorkbook) workbook).getProperties();
                POIXMLProperties.CoreProperties coreProps = xmlProps.getCoreProperties();
                coreProps.setCreator(author);
            }
        }

        return workbook;
    }
}
