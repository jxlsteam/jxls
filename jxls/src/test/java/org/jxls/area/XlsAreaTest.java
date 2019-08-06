package org.jxls.area;

import static org.junit.Assert.assertEquals;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.jxls.command.Command;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;
import org.jxls.transform.Transformer;
import org.mockito.Mock;
import org.mockito.junit.MockitoJUnitRunner;

@RunWith(MockitoJUnitRunner.class)
public class XlsAreaTest {
    @Mock
    private Transformer transformer;
    @Mock
    private Command command1;
    @Mock
    private Command command2;

    @Test
    public void applyAtToAnotherSheet() throws Exception {
        XlsArea xlsArea = new XlsArea(new CellRef("sheet1!A1"), new CellRef("sheet1!G10"), transformer);
        Context context = new Context();
        Size size = xlsArea.applyAt(new CellRef("sheet2!B2"), context);
        assertEquals("Width is wrong", 7, size.getWidth());
        assertEquals("Height is wrong", 10, size.getHeight());
        verify(transformer).transform(new CellRef("sheet1!A1"), new CellRef("sheet2!B2"), context, false);
        verify(transformer).transform(new CellRef("sheet1!D2"), new CellRef("sheet2!E3"), context, false);
    }

    @Test
    public void applyAtWithOneCommand() throws Exception {
        XlsArea xlsArea = new XlsArea(new CellRef("sheet1!A1"), new CellRef("sheet1!G10"), transformer);
        xlsArea.addCommand("sheet1!B3:C5", command1);
        Context context = new Context();
        when(command1.applyAt(new CellRef("sheet2!C4"), context)).thenReturn(new Size(3, 4));
        Size size = xlsArea.applyAt(new CellRef("sheet2!B2"), context);
        assertEquals("Width is wrong", 8, size.getWidth());
        assertEquals("Height is wrong", 11, size.getHeight());
        verify(transformer).transform(new CellRef("sheet1!B6"), new CellRef("sheet2!C8"), context, false);
        verify(transformer).transform(new CellRef("sheet1!D4"), new CellRef("sheet2!F5"), context, false);
    }

    @Test
    public void applyAtShiftDownWithTwoCommands() throws Exception {
        XlsArea xlsArea = new XlsArea(new CellRef("sheet1!A1"), new CellRef("sheet1!G10"), transformer);
        xlsArea.addCommand("sheet1!B3:C5", command1);
        xlsArea.addCommand("sheet1!A7:B8", command2);
        Context context = new Context();
        when(command1.applyAt(new CellRef("sheet2!C4"), context)).thenReturn(new Size(3, 4));
        when(command2.applyAt(new CellRef("sheet2!B9"), context)).thenReturn(new Size(3, 3));
        Size size = xlsArea.applyAt(new CellRef("sheet2!B2"), context);
        assertEquals("Width is wrong", 8, size.getWidth());
        assertEquals("Height is wrong", 12, size.getHeight());
        verify(transformer).transform(new CellRef("sheet1!B6"), new CellRef("sheet2!C8"), context, false);
        verify(transformer).transform(new CellRef("sheet1!D4"), new CellRef("sheet2!F5"), context, false);
        verify(transformer).transform(new CellRef("sheet1!A9"), new CellRef("sheet2!B12"), context, false);
        verify(transformer).transform(new CellRef("sheet1!B9"), new CellRef("sheet2!C12"), context, false);
        verify(transformer).transform(new CellRef("sheet1!C7"), new CellRef("sheet2!E9"), context, false);
        verify(transformer).transform(new CellRef("sheet1!C9"), new CellRef("sheet2!D11"), context, false);
    }

    @Test
    public void applyAtShiftUpWithTwoCommands() throws Exception {
        XlsArea xlsArea = new XlsArea(new CellRef("sheet1!A1"), new CellRef("sheet1!G10"), transformer);
        xlsArea.addCommand("sheet1!B3:C5", command1);
        xlsArea.addCommand("sheet1!A7:B8", command2);
        Context context = new Context();
        when(command1.applyAt(new CellRef("sheet2!C4"), context)).thenReturn(new Size(2, 2));
        when(command2.applyAt(new CellRef("sheet2!B8"), context)).thenReturn(new Size(2, 2));
        when(command1.getLockRange()).thenReturn(true);
        when(command2.getLockRange()).thenReturn(true);
        Size size = xlsArea.applyAt(new CellRef("sheet2!B2"), context);
        assertEquals("Width is wrong", 7, size.getWidth());
        assertEquals("Height is wrong", 10, size.getHeight());
        verify(transformer).transform(new CellRef("sheet1!B6"), new CellRef("sheet2!C6"), context, false);
        verify(transformer).transform(new CellRef("sheet1!A6"), new CellRef("sheet2!B7"), context, false);
        verify(transformer).transform(new CellRef("sheet1!D4"), new CellRef("sheet2!E5"), context, false);
        verify(transformer).transform(new CellRef("sheet1!A9"), new CellRef("sheet2!B10"), context, false);
        verify(transformer).transform(new CellRef("sheet1!B9"), new CellRef("sheet2!C10"), context, false);
        verify(transformer).transform(new CellRef("sheet1!C7"), new CellRef("sheet2!D7"), context, false);
        verify(transformer).transform(new CellRef("sheet1!C10"), new CellRef("sheet2!D10"), context, false);
    }
}
