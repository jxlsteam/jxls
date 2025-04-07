package org.jxls.transform.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.jxls.builder.xls.XlsCommentAreaBuilder;

/**
 * POI utility methods
 * 
 * @author Leonid Vysochyn
 */
public class PoiUtil {
    
    public static void setCellComment(Cell cell, String commentText, String commentAuthor, ClientAnchor anchor) {
        Sheet sheet = cell.getSheet();
        Workbook wb = sheet.getWorkbook();
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        CreationHelper factory = wb.getCreationHelper();
        if (anchor == null) {
            anchor = createAnchor(cell, 1, 3, 0, 2);
        }
        Comment comment = drawing.createCellComment(anchor);
        comment.setString(factory.createRichTextString(commentText));
        comment.setAuthor(commentAuthor != null ? commentAuthor : "");
        cell.setCellComment(comment);
    }
    
    public static ClientAnchor createAnchor(Cell cell, int addToCol1, int addToCol2, int addToRow1, int addToRow2) {
        Sheet sheet = cell.getSheet();
        Workbook wb = sheet.getWorkbook();
        CreationHelper factory = wb.getCreationHelper();
        ClientAnchor anchor = factory.createClientAnchor();
        anchor.setCol1(cell.getColumnIndex() + addToCol1);
        anchor.setCol2(cell.getColumnIndex() + addToCol2);
        anchor.setRow1(cell.getRowIndex() + addToRow1);
        anchor.setRow2(cell.getRowIndex() + addToRow2);
        return anchor;
    }

    public WritableCellValue hyperlink(String address, String link, String linkTypeString) {
        return new WritableHyperlink(address, link, linkTypeString);
    }

    public WritableCellValue hyperlink(String address, String title) {
        return new WritableHyperlink(address, title);
    }

    public static boolean isJxComment(String cellComment) {
        if (cellComment == null) return false;
        String[] commentLines = cellComment.split("\\n");
        for (String commentLine : commentLines) {
            if ((commentLine != null) && XlsCommentAreaBuilder.isCommandString(commentLine.trim())) {
                return true;
            } else if (commentLine.trim().endsWith(">") || commentLine.trim().startsWith("<")) {
            	return true;
            }
        }
        return false;
    }
}
