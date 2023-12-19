package org.jxls.util;

import static java.util.Arrays.asList;
import static java.util.Collections.singletonList;
import static org.junit.Assert.assertEquals;

import java.util.List;

import org.junit.Test;
import org.jxls.area.Area;
import org.jxls.area.XlsArea;
import org.jxls.builder.JxlsTemplateFiller;
import org.jxls.command.EachCommand;
import org.jxls.common.AreaRef;
import org.jxls.common.CellRef;
import org.jxls.common.Size;
import org.jxls.formula.AbstractFormulaProcessor;

public class UtilTest {
    // see more tests in CreateTargetCellRefTest

    @Test
    public void should_return_sheet_names_of_multi_sheet_template() {
        // GIVEN
        EachCommand eachCommand = new EachCommand();
        eachCommand.setMultisheet("multiSheetOutputNames");

        Area areaWithMultiSheetOutputCommand = new XlsArea(new CellRef("areaWithMultiSheetOutput", 1, 1), new Size(1, 1));
        AreaRef ref = new AreaRef(new CellRef("areaWithMultiSheetOutput", 1, 1), new CellRef("areaWithMultiSheetOutput", 1, 1));
        areaWithMultiSheetOutputCommand.addCommand(ref, eachCommand);

        Area areaWithoutMultiSheetOutputCommand = new XlsArea(new CellRef("areaWithoutMultiSheetOutput", 1, 1), new Size());

        // WHEN
        List<String> sheetsNameOfMultiSheetTemplate = JxlsTemplateFiller.getSheetsNameOfMultiSheetTemplate(
                asList(areaWithMultiSheetOutputCommand, areaWithoutMultiSheetOutputCommand));

        // THEN
        assertEquals(sheetsNameOfMultiSheetTemplate, singletonList("areaWithMultiSheetOutput"));
    }
    
// #240 not part of v2.13.0    @Test
    public void test_getFormulaCellRefs_tableSyntax() {
        // Test
        String table = "_tabu";
        String columnHeader = " 1 column b";
        List<String> formulaCellRefs = AbstractFormulaProcessor.getFormulaCellRefs(table + "[" + columnHeader + "]");
        
        // Verify
        assertEquals(1, formulaCellRefs.size());
        assertEquals("_tabu[ 1 column b]", formulaCellRefs.get(0));
    }

// #240 not part of v2.13.0    @Test
    public void test_getFormulaCellRefs_tableSyntax2() {
        // Test
        String formula = "FUNC(one[AB CD],two[123], -1, three[c])";
        List<String> formulaCellRefs = AbstractFormulaProcessor.getFormulaCellRefs(formula);
        
        // Verify
        assertEquals(3, formulaCellRefs.size());
        assertEquals("one[AB CD]", formulaCellRefs.get(0));
        assertEquals("two[123]", formulaCellRefs.get(1));
        assertEquals("three[c]", formulaCellRefs.get(2));
    }
}
