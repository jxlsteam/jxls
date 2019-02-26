package org.jxls.demo;

import org.jxls.area.XlsArea;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.util.TransformerFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * @author Leonid Vysochyn
 *         Date: 2/9/12
 */
public class FormulaExportDemo {
    static Logger logger = LoggerFactory.getLogger(FormulaExportDemo.class);
    private static String template = "formulas_demo.xls";
    private static String output = "target/formulas_demo_output.xls";

    public static void main(String[] args) throws IOException {
        logger.info("Running Formulas demo");
        execute();
    }

    public static void execute() throws IOException {
        logger.info("Opening input stream");
        try(InputStream is = FormulaExportDemo.class.getResourceAsStream(template)) {
            try (OutputStream os = new FileOutputStream(output)) {
                Transformer transformer = TransformerFactory.createTransformer(is, os);
                XlsArea sheet1Area = new XlsArea("Sheet1!A1:D4", transformer);
                XlsArea sheet2Area = new XlsArea("Sheet2!A1:A2", transformer);
                XlsArea sheet3Area = new XlsArea("'Sheet 3'!A1:A2", transformer);
                Context context = new Context();
                sheet3Area.applyAt(new CellRef("Sheet1!K1"), context);
                sheet2Area.applyAt(new CellRef("Sheet2!B6"), context);
                sheet2Area.applyAt(new CellRef("Sheet2!C6"), context);
                sheet2Area.applyAt(new CellRef("Sheet2!D6"), context);
                sheet1Area.applyAt(new CellRef("Sheet1!F11"), context);
                sheet1Area.processFormulas();
                transformer.write();
                logger.info("written to file");
            }
        }
    }

}
