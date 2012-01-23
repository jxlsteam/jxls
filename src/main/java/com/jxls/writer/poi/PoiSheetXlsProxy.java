package com.jxls.writer.poi;

import com.jxls.writer.Pos;
import com.jxls.writer.XlsProxy;
import com.jxls.writer.command.Context;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;

/**
 * Date: Apr 15, 2009
 *
 * @author Leonid Vysochyn
 */
public class PoiSheetXlsProxy implements XlsProxy {

    Sheet sheet;
    XlsSheetInfo sheetInfo;

    public PoiSheetXlsProxy(Sheet sheet, XlsSheetInfo sheetInfo) {
        this.sheet = sheet;
        this.sheetInfo = sheetInfo;
    }

    public void processCell(Pos pos, Pos newPos, Context context) {
        XlsCellInfo cellInfo = sheetInfo.getPos(pos);
        Row row = sheet.getRow( newPos.getX() );
        if( row == null ){
            row = sheet.createRow( newPos.getX() );
        }
        Cell cell = row.getCell( newPos.getY() );
        if( cell == null ){
            cell = row.createCell( newPos.getY() );
        }
        cellInfo.writeToCell( cell );
    }


}
