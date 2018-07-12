import com.gather.excelcommons.builder.ExcelBuilder;
import com.gather.excelcommons.sheet.DefaultSheetBuilder;
import com.gather.excelcommons.sheet.ISheetBuilder;
import com.gather.excelcommons.sheet.creator.DefaultSheetCreator;
import com.gather.excelcommons.sheet.populator.CompleteSheetPopulator;
import com.gather.excelcommons.sheet.populator.DefaultBodySheetPopulator;
import com.gather.excelcommons.sheet.populator.DefaultHeaderSheetPopulator;
import com.gather.excelcommons.workbook.DefaultWorkbookCreator;
import com.gather.gathercommons.model.DefaultDataTableModel;
import com.gather.gathercommons.model.IDataTableModel;
import org.apache.log4j.Logger;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Created with IntelliJ IDEA.
 * $ Project: excelcommons
 * User: rodrigotroy
 * Date: 10/23/17
 * Time: 16:08
 */
public class ExcelBuilderTest {
    private static final Logger LOG = Logger.getLogger(ExcelBuilderTest.class);
    private static final String RESULT = "test.xlsx";

    private IDataTableModel getDataTableModel() {
        final IDataTableModel dataTableModel = new DefaultDataTableModel();


        List<List<Object>> titles = new ArrayList<>();
        List<Object> title = new ArrayList<>();
        title.add(0);
        title.add("Titulo Tabla");
        title.add("Titulo Tabla");

        titles.add(title);
        dataTableModel.setTitles(titles);

        List<List<Object>> headers = new ArrayList<>();
        List<Object> header1 = new ArrayList<>();
        header1.add("col1");
        header1.add(2);
        header1.add(0);
        header1.add(1);
        header1.add(2);
        header1.add(0);

        headers.add(header1);
        dataTableModel.setHeaders(headers);

        List<List<Object>> rows = new ArrayList<>();
        List<Object> row = new ArrayList<>();
        row.add(1);

        rows.add(row);
        dataTableModel.setRows(rows);

        return dataTableModel;
    }

    private void init() throws
                        Exception {
        List<ISheetBuilder> sheetBuilders = new ArrayList<>();
        final ISheetBuilder sheetBuilder1 = new DefaultSheetBuilder(new DefaultSheetCreator(this.getDataTableModel().getTitles().get(0).get(2)),
                                                                    new CompleteSheetPopulator(new DefaultHeaderSheetPopulator(this.getDataTableModel(),
                                                                                                                               0),
                                                                                               new DefaultBodySheetPopulator(this.getDataTableModel(),
                                                                                                                             1)));
        sheetBuilders.add(sheetBuilder1);

        ExcelBuilder excelBuilder = new ExcelBuilder(new DefaultWorkbookCreator(),
                                                     sheetBuilders);
        excelBuilder.createExcel();

        File file = new File(RESULT);
        FileOutputStream fop = new FileOutputStream(file);

        byte[] contentInBytes = excelBuilder.getStream().toByteArray();

        fop.write(contentInBytes);
        fop.flush();
        fop.close();

    }

    public static void main(String[] args) throws
                                           Exception {
        LOG.info("INICIO TEST");

        ExcelBuilderTest excelBuilderTest = new ExcelBuilderTest();
        excelBuilderTest.init();

        LOG.info("FIN TEST");
    }
}
