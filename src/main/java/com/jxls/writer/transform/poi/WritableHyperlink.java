package com.jxls.writer.transform.poi;

import com.jxls.writer.common.Context;
import org.apache.poi.ss.usermodel.*;

/**
 * @author Leonid Vysochyn
 *         Date: 6/18/12 4:01 PM
 */
public class WritableHyperlink implements WritableCellValue {
    public static final String LINK_URL = "URL";
    public static final String LINK_DOCUMENT = "DOCUMENT";
    public static final String LINK_EMAIL = "EMAIL";
    public static final String LINK_FILE = "FILE";

    String address;
    String title;
    int linkType;

    CellStyle linkStyle;

    public WritableHyperlink(String address, String linkTypeString) {
        this(address, null, linkTypeString);
    }

    public WritableHyperlink(String address, String title, String linkTypeString) {
        this.address = address;
        this.title = title;
        if( LINK_URL.equals(linkTypeString) ){
            linkType = Hyperlink.LINK_URL;
        }else if( LINK_DOCUMENT.equals(linkTypeString) ){
            linkType = Hyperlink.LINK_DOCUMENT;
        }else if( LINK_EMAIL.equals(linkTypeString) ){
            linkType = Hyperlink.LINK_EMAIL;
        }else if( LINK_FILE.equals(linkTypeString) ){
            linkType = Hyperlink.LINK_FILE;
        }else {
            throw new IllegalArgumentException("Link type must be one of the following values: " + LINK_URL + "," + LINK_DOCUMENT + "," + LINK_EMAIL + "," + LINK_FILE);
        }
    }

    public WritableHyperlink(String address, String title, int linkType) {
        this.address = address;
        this.title = title;
        this.linkType = linkType;
    }

    public Object writeToCell(Cell cell, Context context) {
        Workbook workbook = cell.getSheet().getWorkbook();
        CreationHelper createHelper = workbook.getCreationHelper();
        Hyperlink hyperlink = createHelper.createHyperlink(linkType);
        hyperlink.setAddress(address);
        cell.setHyperlink(hyperlink);
        cell.setCellValue(title);
        if( linkStyle == null ){
            linkStyle = createLinkStyle(workbook);
        }
        cell.setCellStyle(linkStyle);
        return cell;
    }

    private CellStyle createLinkStyle(Workbook workbook) {
        linkStyle = workbook.createCellStyle();
        Font hlink_font = workbook.createFont();
        hlink_font.setUnderline(Font.U_SINGLE);
        hlink_font.setColor(IndexedColors.BLUE.getIndex());
        linkStyle.setFont(hlink_font);
        return linkStyle;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

}
