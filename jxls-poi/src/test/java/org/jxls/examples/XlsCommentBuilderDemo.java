package org.jxls.examples;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.junit.Assert;
import org.junit.Test;
import org.jxls.Jxls3Tester;
import org.jxls.TestWorkbook;
import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.command.EachCommand;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.ContextImpl;
import org.jxls.entity.Department;
import org.jxls.formula.StandardFormulaProcessor;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.PoiUtil;

/**
 * @author Leonid Vysochyn
 */
public class XlsCommentBuilderDemo {
    private static final String template = "comment_markup_demo.xls";
    private static final String output = "target/comment_builder_output.xls";

    @Test
    public void test() throws IOException {
        List<Department> departments = Department.createDepartments();
        try (InputStream is = XlsCommentBuilderDemo.class.getResourceAsStream(template)) {
            try (OutputStream os = new FileOutputStream(output)) {
                Transformer transformer = Jxls3Tester.createTransformer(is, os);
                AreaBuilder areaBuilder = new XlsCommentAreaBuilder();
                List<Area> xlsAreaList = areaBuilder.build(transformer, false);
                Area xlsArea = xlsAreaList.get(0);
                Context context = new ContextImpl();
                context.putVar("util", new PoiUtil());
                context.putVar("departments", departments);
                xlsArea.applyAt(new CellRef("Down!A1"), context);
                StandardFormulaProcessor fp = new StandardFormulaProcessor();
                xlsArea.processFormulas(fp);
                xlsArea.reset();
                EachCommand eachCommand = (EachCommand) xlsArea.findCommandByName("each").get(0);
                eachCommand.setDirection(EachCommand.Direction.RIGHT);
                xlsArea.applyAt(new CellRef("Right!A1"), context);
                xlsArea.processFormulas(fp);
                transformer.write();
            }
        }
        
        // Verify hyperlink
        try (TestWorkbook w = new TestWorkbook(new File(output))) {
            w.selectSheet("Down");
            Assert.assertEquals("http://jxls.sf.net=http://jxls.sf.net", w.getHyperlink(5, 2));
        }
    }
}
