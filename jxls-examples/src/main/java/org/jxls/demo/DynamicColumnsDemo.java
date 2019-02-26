package org.jxls.demo;

import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xml.XmlAreaBuilder;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.PoiTransformer;
import org.jxls.util.TransformerFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * Created by Leonid Vysochyn on 24-Jun-15.
 */
public class DynamicColumnsDemo {
    static Logger logger = LoggerFactory.getLogger(DynamicColumnsDemo.class);

    private static final String DYNAMIC_COLUMNS_DEMO_XML_CONFIG = "dynamic_columns_demo.xml";
    private static final String TEMPLATE = "dynamic_columns_demo.xls";
    private static final String OUTPUT = "target/dynamic_columns_output.xls";

    public static void main(String[] args) throws IOException {
        logger.info("Running Dynamic Columns demo");
        execute();
    }

    public static void execute() throws IOException {
        logger.info("Preparing test data");
        List<String> headers = new ArrayList<String>(Arrays.asList("Param 1", "Param 2", "Param 3", "Param 4", "Param 5"));
        List<List<Object>> rows = new ArrayList<>();
        for(int i = 0; i < 10; i++){
            List<Object> row = new ArrayList<>();
            for(int k = 0; k < 5; k++){
                row.add(String.format("Val(%s,%s)", i, k));
            }
            rows.add(row);
        }

        // loading areas and commands using XmlAreaBuilder
        try(InputStream is = DynamicColumnsDemo.class.getResourceAsStream(TEMPLATE)) {
            try (OutputStream os = new FileOutputStream(OUTPUT)) {
                Transformer transformer = TransformerFactory.createTransformer(is, os);
                InputStream configInputStream = DynamicColumnsDemo.class.getResourceAsStream(DYNAMIC_COLUMNS_DEMO_XML_CONFIG);
                AreaBuilder areaBuilder = new XmlAreaBuilder(configInputStream, transformer);
                List<Area> xlsAreaList = areaBuilder.build();
                Area xlsArea = xlsAreaList.get(0);
                // creating context
                Context context = PoiTransformer.createInitialContext();
                context.putVar("headers", headers);
                context.putVar("rows", rows);
                // applying transformation
                logger.info("Applying area " + xlsArea.getAreaRef() + " at cell " + new CellRef("Result!A1"));
                xlsArea.applyAt(new CellRef("Result!A1"), context);
                // saving the results to file
                transformer.write();
                logger.info("Complete");
            }
        }
    }
}
