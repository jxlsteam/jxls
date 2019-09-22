package org.jxls.demo.issue;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.transform.poi.PoiContext;
import org.jxls.transform.poi.PoiTransformer;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class Issue160TestCase {

    public static void main(String[] args) throws IOException, ParseException {
        List<Map<String, Object>> lotsOfStuff = createLotsOfStuff();
        Context context = new PoiContext();
        context.putVar("lotsOfStuff", lotsOfStuff);
        context.putVar("columns", new Columns());
        try(InputStream in = Issue160TestCase.class.getResourceAsStream("issue160_template.xlsx")) {
            try (OutputStream os = new FileOutputStream("target/issue160_output.xlsx")) {
                Workbook workbook = WorkbookFactory.create(in);
                PoiTransformer transformer = PoiTransformer.createSxssfTransformer(workbook, 2, false);
                AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer);
                List<Area> xlsAreaList = areaBuilder.build();
                Area xlsArea = xlsAreaList.get(0);
                xlsArea.applyAt(new CellRef("Result!A1"), context);
                SXSSFWorkbook workbook2 = (SXSSFWorkbook) transformer.getWorkbook();
                workbook2.write(os);
            }
        }
    }

    private static List<Map<String, Object>> createLotsOfStuff() {
        Map<String, Object> stuff1 = new LinkedHashMap<>();
        Map<String, Object> stuff2 = new LinkedHashMap<>();

        stuff1.put("header0", "stuff_1_value0");
        stuff1.put("header1_dynamic", "stuff_1_value1");
        stuff1.put("header2_dynamic", "stuff_1_value2");
        stuff1.put("header3_dynamic", "stuff_1_value3");

        stuff2.put("header0", "stuff_2_value0");
        stuff2.put("header1_dynamic", "stuff_2_value1");
        stuff2.put("header2_dynamic", "stuff_2_value2");
        stuff2.put("header3_dynamic", "stuff_2_value3");

        return Arrays.asList(stuff1, stuff2);
    }
}
