package org.jxls.transformer;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFTableColumn;
import org.jxls.common.Context;
import org.jxls.templatebasedtests.PivotTableTest;
import org.jxls.transform.Transformer;
import org.jxls.transform.TransformerDelegator;
import org.jxls.transform.poi.PoiTransformer;

/**
 * @see PivotTableTest
 */
public class TestPTTransformer extends TransformerDelegator {
    private final Context context;

    public TestPTTransformer(Transformer transformer, Context context) {
        super(transformer);
        this.context = context;
    }

    protected void beforeWrite() {
        Workbook workbook = ((PoiTransformer) transformer).getWorkbook();
        for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
            XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(sheetIndex);
            for (XSSFTable table : sheet.getTables()) {
                processTable(table);
            }
        }
    }

    private void processTable(XSSFTable table) {
        for (XSSFTableColumn col : table.getColumns()) {
            col.setName(evaluate(col.getName()));
        }
        table.updateReferences();
        table.updateHeaders();
    }

    private String evaluate(String name) {
        name = name.replace("${", "").replace("}", ""); // XXX Krücke
        return (String) getTransformationConfig().getExpressionEvaluator().evaluate(name, context.toMap());
    }
}
