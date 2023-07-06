package org.jxls.examples;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xml.XmlAreaBuilder;
import org.jxls.command.AbstractCommand;
import org.jxls.command.Command;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;
import org.jxls.entity.Department;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.PoiTransformer;
import org.jxls.util.TransformerFactory;
import org.jxls.util.Util;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author Leonid Vysochyn
 *         Date: 2/21/12 4:30 PM
 */
public class UserCommandDemo {
    private static final Logger logger = LoggerFactory.getLogger(UserCommandDemo.class);
    private static final String template = "each_if_demo.xls";
    private static final String xmlConfig = "user_command_demo.xml";
    private static final String output = "target/user_command_output.xls";

    @Test
    public void test() throws IOException {
        logger.info("Running User Command demo");
        List<Department> departments = Department.createDepartments();
        logger.info("Opening input stream");
        try (InputStream is = EachIfCommandDemo.class.getResourceAsStream(template)) {
            try (OutputStream os = new FileOutputStream(output)) {
                Transformer transformer = TransformerFactory.createTransformer(is, os);
                logger.info("Creating areas");
                try (InputStream configInputStream = UserCommandDemo.class.getResourceAsStream(xmlConfig)) {
                    AreaBuilder areaBuilder = new XmlAreaBuilder(configInputStream, transformer);
                    List<Area> xlsAreaList = areaBuilder.build();
                    Area xlsArea = xlsAreaList.get(0);
                    Context context = new Context();
                    context.putVar("departments", departments);
                    logger.info("Applying area at cell " + new CellRef("Down!A1"));
                    xlsArea.applyAt(new CellRef("Down!A1"), context);
                    xlsArea.processFormulas();
                    logger.info("Complete");
                    transformer.write();
                    logger.info("Written to file");
                }
            }
        }
    }
    
    // also used by UserCommandExcelMarkupDemo
    /**
     * An implementation of a Command for row grouping
     */
    public static class GroupRowCommand extends AbstractCommand {
        private Area area;
        private String collapseIf;

        @Override
        public String getName() {
            return "groupRow";
        }

        @Override
        public Size applyAt(CellRef cellRef, Context context) {
            Size resultSize = area.applyAt(cellRef, context);
            if (resultSize.equals(Size.ZERO_SIZE)) {
                return resultSize;
            }
            PoiTransformer transformer = (PoiTransformer) area.getTransformer();
            Workbook workbook = transformer.getWorkbook();
            Sheet sheet = workbook.getSheet(cellRef.getSheetName());
            int startRow = cellRef.getRow();
            int endRow = cellRef.getRow() + resultSize.getHeight() - 1;
            sheet.groupRow(startRow, endRow);
            if (collapseIf != null && collapseIf.trim().length() > 0) {
                boolean collapseFlag = Util.isConditionTrue(getTransformationConfig().getExpressionEvaluator(), collapseIf, context);
                sheet.setRowGroupCollapsed(startRow, collapseFlag);
            }
            return resultSize;
        }

        @Override
        public Command addArea(Area area) {
            super.addArea(area);
            this.area = area;
            return this;
        }

        public void setCollapseIf(String collapseIf) {
            this.collapseIf = collapseIf;
        }
    }
}
