package org.jxls.common;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.util.List;

import org.junit.Assert;
import org.junit.Test;

public class CellDataJTest {

    @Test
    public void sheetname_col_row() {
        CellData cellData = new CellData("sheet1", 2, 3);
        
        assertEquals("sheet1", cellData.getSheetName());
        assertEquals(2, cellData.getRow());
        assertEquals(3, cellData.getCol());
    }
    
    @Test
    public void equality() {
        Assert.assertTrue (new CellData("sheet1", 5, 10, CellData.CellType.STRING, "Abc")
                   .equals(new CellData("sheet1", 5, 10, CellData.CellType.STRING, "Abc")));
        Assert.assertFalse(new CellData("sheet1", 5, 10, CellData.CellType.STRING, "Abc1")
                   .equals(new CellData("sheet1", 5, 10, CellData.CellType.STRING, "Abc")));
    }
    
    @Test
    public void creation_with_sheetName_row_col_type_value() {
        CellData cellData = new CellData("Sheet1", 5, 10, CellData.CellType.NUMBER, 1.2);

        assertEquals("Sheet1", cellData.getSheetName());
        assertEquals(5, cellData.getRow());
        assertEquals(10, cellData.getCol());
        assertEquals(CellData.CellType.NUMBER, cellData.getCellType());
        assertEquals(1.2, cellData.getCellValue());
        assertEquals(new CellRef("Sheet1", 5, 10), cellData.getCellRef());
    }

    @Test
    public void creation_with_pos_type_value() {
        CellData cellData = new CellData(new CellRef("Sheet1", 5, 10), CellData.CellType.STRING, "Abc");

        assertEquals(new CellRef("Sheet1", 5, 10), cellData.getCellRef());
        assertEquals("Sheet1", cellData.getSheetName());
        assertEquals(5, cellData.getRow());
        assertEquals(10, cellData.getCol());
    }

    @Test
    public void add_target_pos() {
        CellData cellData = new CellData("sheet1", 2, 3);
        cellData.addTargetPos(new CellRef("sheet1", 2, 3));
        cellData.addTargetPos(new CellRef("sheet2", 3, 4));
        List<CellRef> targetPos = cellData.getTargetPos();

        assertEquals(targetPos.size(), 2);
        assertTrue(targetPos.contains(new CellRef("sheet1", 2, 3)));
        assertTrue(targetPos.contains(new CellRef("sheet2", 3, 4)));
    }

    @Test
    public void get_pos() {
        CellData cellData = new CellData("sheet1", 2, 3);

        assertEquals(cellData.getCellRef(), new CellRef("sheet1", 2, 3));
    }
    
    @Test
    public void reset_target_pos() {
        CellData cellData = new CellData("sheet1", 1, 2, CellData.CellType.STRING, "Abc");
        cellData.addTargetPos(new CellRef("sheet2", 0, 0));
        cellData.addTargetPos(new CellRef("sheet1", 1, 1));
        cellData.resetTargetPos();

        assertTrue(cellData.getTargetPos().isEmpty());
    }

    @Test
    public void set_get_comment() {
        CellData cellData = new CellData("sheet1", 1, 2, CellData.CellType.STRING, "Abc");
        cellData.setCellComment("Test comment");

        assertEquals(cellData.getCellComment(), "Test comment");
    }

}
