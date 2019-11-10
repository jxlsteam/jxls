package org.jxls.unittests;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;

import java.util.ArrayList;
import java.util.List;

import org.junit.Assert;
import org.junit.Test;
import org.jxls.command.SheetNameGenerator;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.templatebasedtests.multisheet.PoiSafeSheetNameBuilderTest;
import org.jxls.transform.SafeSheetNameBuilder;
import org.jxls.transform.poi.PoiContext;
import org.jxls.transform.poi.PoiSafeSheetNameBuilder;

/**
 * This is a multi sheet test.
 * 
 * @see PoiSafeSheetNameBuilderTest
 */
public class PoiSafeSheetNameBuilderUnitTest {
    
    /**
     * Tests PoiSafeSheetNameBuilder.
     */
    @Test
    public void testSafeSheetNames() {
        Context context = new PoiContext();

        List<String> sheetNames = new ArrayList<>();
        sheetNames.add("sheet 1");
        sheetNames.add("sheet 2");
        sheetNames.add("sheet 3 []_=' - this is a very long name with special characters");
        SheetNameGenerator generator = new SheetNameGenerator(sheetNames, new CellRef("A1"));

        assertEquals("'sheet 1'!A1", generator.generateCellRef(0, context).toString());
        assertEquals("'sheet 2'!A1", generator.generateCellRef(1, context).toString());
        assertEquals("Name contains invalid chars and/or is too long",
                "'sheet 3   _='' - this is a very '!A1", generator.generateCellRef(2, context).toString());
    }

    /**
     * Sheet names must be unique.
     */
    @Test
    public void testUniqueSheetNames_simple() {
        Context context = new PoiContext();
        Assert.assertTrue("precondition: PoiContext must contain " + PoiSafeSheetNameBuilder.class.getSimpleName(),
                context.getVar(SafeSheetNameBuilder.CONTEXT_VAR_NAME) instanceof PoiSafeSheetNameBuilder);

        List<String> sheetNames = new ArrayList<>();
        sheetNames.add("a");
        sheetNames.add("a");
        sheetNames.add("a");
        SheetNameGenerator generator = new SheetNameGenerator(sheetNames, new CellRef("A1"));

        assertEquals("1st sheet name not okay", "a", generator.generateCellRef(0, context).getSheetName());
        assertEquals("2nd sheet name not okay", "a(1)", generator.generateCellRef(1, context).getSheetName());
        assertEquals("3rd sheet name not okay", "a(2)", generator.generateCellRef(2, context).getSheetName());
    }

    /**
     * Sheet names must be unique.
     * In this testcase the given sheet names contain an invalid char and are too long. And there's also a collision (...-4).
     */
    @Test
    public void testUniqueSheetNames() {
        Context context = new Context();
        context.putVar(SafeSheetNameBuilder.CONTEXT_VAR_NAME, new PoiSafeSheetNameBuilder() {
            @Override
            protected int getFirstSerialNumber() {
                return 2;
            }
            
            @Override
            protected String addSerialNumber(String text, int serialNumber) {
                return text + "-" + serialNumber;
            }
        });

        List<String> sheetNames = new ArrayList<>();
        for (int i = 1; i <= 10; i++) {
            sheetNames.add("a[aaaaaaaaaaaaaaaaaaaaaaaaaaaaax"); // text has invalid char and is too long
        }
        sheetNames.add("a aaaaaaaaaaaaaaaaaaaaaaaaaaa-4"); // becomes "...-11", that's okay
        sheetNames.add("b"); // no change
        SheetNameGenerator generator = new SheetNameGenerator(sheetNames, new CellRef("A1"));

        assertEquals("1st sheet name not okay", "a aaaaaaaaaaaaaaaaaaaaaaaaaaaaa", generator.generateCellRef(0, context).getSheetName());
        assertEquals("2nd sheet name not okay", "a aaaaaaaaaaaaaaaaaaaaaaaaaaa-2", generator.generateCellRef(1, context).getSheetName());
        assertEquals("3rd sheet name not okay", "a aaaaaaaaaaaaaaaaaaaaaaaaaaa-3", generator.generateCellRef(2, context).getSheetName());
        assertEquals("4th sheet name not okay", "a aaaaaaaaaaaaaaaaaaaaaaaaaaa-4", generator.generateCellRef(3, context).getSheetName());
        assertEquals("5th sheet name not okay", "a aaaaaaaaaaaaaaaaaaaaaaaaaaa-5", generator.generateCellRef(4, context).getSheetName());
        assertEquals("6th sheet name not okay", "a aaaaaaaaaaaaaaaaaaaaaaaaaaa-6", generator.generateCellRef(5, context).getSheetName());
        assertEquals("7th sheet name not okay", "a aaaaaaaaaaaaaaaaaaaaaaaaaaa-7", generator.generateCellRef(6, context).getSheetName());
        assertEquals("8th sheet name not okay", "a aaaaaaaaaaaaaaaaaaaaaaaaaaa-8", generator.generateCellRef(7, context).getSheetName());
        assertEquals("9th sheet name not okay", "a aaaaaaaaaaaaaaaaaaaaaaaaaaa-9", generator.generateCellRef(8, context).getSheetName());
        assertEquals("10th sheet name not okay","a aaaaaaaaaaaaaaaaaaaaaaaaaa-10", generator.generateCellRef(9, context).getSheetName());
        assertEquals("11th sheet name not okay","a aaaaaaaaaaaaaaaaaaaaaaaaaa-11", generator.generateCellRef(10, context).getSheetName());
        assertEquals("12th sheet name not okay", "b", generator.generateCellRef(11, context).getSheetName());
    }

    /**
     * sheetNames array has only 2 entries. However, there are 4 sheets.
     */
    @Test
    public void testNotEnoughSheetNames() {
        Context context = new Context();
        context.putVar(SafeSheetNameBuilder.CONTEXT_VAR_NAME, new PoiSafeSheetNameBuilder() {
            @Override
            public String createSafeSheetName(String givenSheetName, int index) {
                if (givenSheetName == null) {
                    givenSheetName = "sheet " + (index + 1);
                }
                return super.createSafeSheetName(givenSheetName, index);
            }
        });

        List<String> sheetNames = new ArrayList<>();
        sheetNames.add("first");
        sheetNames.add("2nd");
        SheetNameGenerator generator = new SheetNameGenerator(sheetNames, new CellRef("A1"));

        assertEquals("1st sheet name not okay", "first", generator.generateCellRef(0, context).getSheetName());
        assertEquals("2nd sheet name not okay", "2nd", generator.generateCellRef(1, context).getSheetName());
        assertEquals("3rd sheet name not okay", "sheet 3", generator.generateCellRef(2, context).getSheetName());
        assertEquals("4th sheet name not okay", "sheet 4", generator.generateCellRef(3, context).getSheetName());
    }

    /**
     * Only two sheet names given. However, there are 4 sheets. Test this case with the default behaviour.
     */
    @Test
    public void testNotEnoughSheetNames_noSafeSheetNameBuilder() {
        Context context = new Context();
        Assert.assertNull("precondition: Context must not contain " + PoiSafeSheetNameBuilder.class.getSimpleName(),
                context.getVar(SafeSheetNameBuilder.CONTEXT_VAR_NAME));

        List<String> sheetNames = new ArrayList<>();
        sheetNames.add("first");
        SheetNameGenerator generator = new SheetNameGenerator(sheetNames, new CellRef("A1"));

        assertEquals("first", generator.generateCellRef(0, context).getSheetName());
        assertNull(generator.generateCellRef(1, context));
    }

    @Test
    public void testSheetNamesWithSerialNumber() {
        Context context = new Context();
        context.putVar(SafeSheetNameBuilder.CONTEXT_VAR_NAME, new PoiSafeSheetNameBuilder() {
            @Override
            public String createSafeSheetName(String givenSheetName, int index) {
                return super.createSafeSheetName((index + 1) + ". " + givenSheetName, index);
            }
        });

        List<String> sheetNames = new ArrayList<>();
        sheetNames.add("Finanzen");
        sheetNames.add("Rechnungen");
        sheetNames.add("Belege");
        sheetNames.add("Finanzen");
        SheetNameGenerator generator = new SheetNameGenerator(sheetNames, new CellRef("A1"));

        assertEquals("1. Finanzen", generator.generateCellRef(0, context).getSheetName());
        assertEquals("2. Rechnungen", generator.generateCellRef(1, context).getSheetName());
        assertEquals("3. Belege", generator.generateCellRef(2, context).getSheetName());
        assertEquals("4. Finanzen", generator.generateCellRef(3, context).getSheetName());
    }
}
