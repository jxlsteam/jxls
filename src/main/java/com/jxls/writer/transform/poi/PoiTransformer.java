package com.jxls.writer.transform.poi;

import com.jxls.writer.Pos;
import com.jxls.writer.command.Context;
import com.jxls.writer.transform.Transformer;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author Leonid Vysochyn
 *         Date: 1/23/12 2:36 PM
 */
public class PoiTransformer implements Transformer {

    Workbook workbook;

    public PoiTransformer(Workbook workbook) {
        this.workbook = workbook;
    }

    @Override
    public void transform(Pos pos, Pos newPos, Context context) {
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(pos.getY());
        Cell cell = row.getCell(pos.getX());
        if (cell != null) {
            CellData cellData = CellData.createCellData(cell);
            Row destRow = sheet.getRow(newPos.getY());
            if (destRow == null) {
                destRow = sheet.createRow(newPos.getY());
            }
            Cell destCell = destRow.getCell(newPos.getX());
            if (destCell == null) {
                destCell = destRow.createCell(newPos.getX());
            }
            cellData.writeToCell(destCell, context);
        }
    }
}
