package com.jxls.writer.transform.poi;

import org.apache.poi.ss.usermodel.*;

/**
 * @author Leonid Vysochyn
 */
public class PoiUtil {
    public static void setCellComment(Cell cell, String commentText, String commentAuthor, ClientAnchor anchor){
        Sheet sheet = cell.getSheet();
        Workbook wb = sheet.getWorkbook();
        Drawing drawing = sheet.createDrawingPatriarch();
        CreationHelper factory = wb.getCreationHelper();
        if( anchor == null ){
            anchor = factory.createClientAnchor();
            anchor.setCol1(0);
            anchor.setCol2(1);
            anchor.setRow1(0);
            anchor.setRow2(1);
        }
        Comment comment = drawing.createCellComment(anchor);
        comment.setString(factory.createRichTextString(commentText));
        comment.setAuthor(commentAuthor != null ? commentAuthor : "");
        cell.setCellComment( comment );
    }

    public WritableCellValue hyperlink(String address, String link, String linkTypeString){
        return new WritableHyperlink(address, link, linkTypeString);
    }

    public WritableCellValue hyperlink(String address, String linkTypeString){
        return new WritableHyperlink(address, linkTypeString);
    }
}
