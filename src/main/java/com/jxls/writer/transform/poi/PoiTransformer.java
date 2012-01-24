package com.jxls.writer.transform.poi;

import com.jxls.writer.Cell;
import com.jxls.writer.command.Context;
import com.jxls.writer.transform.Transformer;
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
    public void transform(Cell pos, Cell newCell, Context context) {
        Sheet sheet = workbook.getSheetAt(pos.getSheetIndex());
        Row row = sheet.getRow(pos.getRow());
        org.apache.poi.ss.usermodel.Cell cell = row.getCell(pos.getCol());
        if (cell != null) {
            CellData cellData = CellData.createCellData(cell);
            Sheet destSheet = workbook.getSheetAt(newCell.getSheetIndex());
            Row destRow = destSheet.getRow(newCell.getRow());
            if (destRow == null) {
                destRow = destSheet.createRow(newCell.getRow());
            }
            org.apache.poi.ss.usermodel.Cell destCell = destRow.getCell(newCell.getCol());
            if (destCell == null) {
                destCell = destRow.createCell(newCell.getCol());
            }
            cellData.writeToCell(destCell, context);
        }
    }
}
