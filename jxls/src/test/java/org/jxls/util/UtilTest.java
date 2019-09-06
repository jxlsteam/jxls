package org.jxls.util;

import org.junit.Test;
import org.jxls.area.Area;
import org.jxls.area.XlsArea;
import org.jxls.command.EachCommand;
import org.jxls.common.AreaRef;
import org.jxls.common.CellRef;
import org.jxls.common.Size;

import java.util.List;

import static java.util.Arrays.asList;
import static java.util.Collections.singletonList;
import static org.junit.Assert.assertEquals;
import static org.jxls.util.Util.getSheetsNameOfMultiSheetTemplate;


public class UtilTest {

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
        List<String> sheetsNameOfMultiSheetTemplate = getSheetsNameOfMultiSheetTemplate(asList(areaWithMultiSheetOutputCommand, areaWithoutMultiSheetOutputCommand));

        // THEN
        assertEquals(sheetsNameOfMultiSheetTemplate, singletonList("areaWithMultiSheetOutput"));
    }
}
