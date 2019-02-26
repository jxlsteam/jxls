package org.jxls.demo.guide;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.PoiTransformer;
import org.jxls.util.JxlsHelper;
import org.jxls.util.TransformerFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

/**
 * @author Leonid Vysochyn
 *         Date: 10/22/13
 */
public class HighlightDemo {
    static Logger logger = LoggerFactory.getLogger(HighlightDemo.class);

    public static void main(String[] args) throws ParseException, IOException, InvalidFormatException {
        logger.info("Running Highlight demo");
        List<Employee> employees = generateSampleEmployeeData();
        try(InputStream is = HighlightDemo.class.getResourceAsStream("highlight_template.xls")) {
            try (OutputStream os = new FileOutputStream("target/highlight_output.xls")) {
                PoiTransformer transformer = PoiTransformer.createTransformer(is, os);
                AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer, false);
                List<Area> xlsAreaList = areaBuilder.build();
                Area mainArea = xlsAreaList.get(0);
                Area loopArea = xlsAreaList.get(0).getCommandDataList().get(0).getCommand().getAreaList().get(0);
                loopArea.addAreaListener(new HighlightCellAreaListener(transformer));
                Context context = new Context();
                context.putVar("employees", employees);
                mainArea.applyAt(new CellRef("Result!A1"), context);
                mainArea.processFormulas();
                transformer.write();
            }
        }
    }

    private static List<Employee> generateSampleEmployeeData() throws ParseException {
        List<Employee> employees = new ArrayList<Employee>();
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MMM-dd", Locale.US);
        employees.add( new Employee("Elsa", dateFormat.parse("1970-Jul-10"), 1500, 0.15) );
        employees.add( new Employee("Oleg", dateFormat.parse("1973-Apr-30"), 2300, 0.25) );
        employees.add( new Employee("Neil", dateFormat.parse("1975-Oct-05"), 2500, 0.00) );
        employees.add( new Employee("Maria", dateFormat.parse("1978-Jan-07"), 1700, 0.15) );
        employees.add( new Employee("John", dateFormat.parse("1969-May-30"), 2800, 0.20) );
        return employees;
    }
}
