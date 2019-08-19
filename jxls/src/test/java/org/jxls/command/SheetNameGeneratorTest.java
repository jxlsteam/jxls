package org.jxls.command;

import static org.junit.Assert.assertEquals;

import java.util.ArrayList;
import java.util.List;
import org.junit.Test;
import org.jxls.common.CellRef;
import org.jxls.common.Context;

public class SheetNameGeneratorTest {
    @Test
    public void testMainCase() {
        List<String> sheetNames = new ArrayList<>();

        sheetNames.add("sheet 1");
        sheetNames.add("sheet 2");
        sheetNames.add("sheet 3 []_=' - this is a very long name with special characters");

        SheetNameGenerator generator = new SheetNameGenerator(sheetNames, new CellRef("A1"));
        Context context = new Context();

        assertEquals("'sheet 1'!A1", generator.generateCellRef(0, context).toString());
        assertEquals("'sheet 2'!A1", generator.generateCellRef(1, context).toString());
        assertEquals("'sheet 3   _='' - this is a very '!A1", generator.generateCellRef(2, context).toString());
    }
}
