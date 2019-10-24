package org.jxls.transform.poi;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Workbook;
import org.jxls.common.Context;

/**
 * Writable cell value implementation for Hyperlink
 * @author Leonid Vysochyn
 */
public class WritableHyperlink implements WritableCellValue {
    public static final String LINK_URL = "URL";
    public static final String LINK_DOCUMENT = "DOCUMENT";
    public static final String LINK_EMAIL = "EMAIL";
    public static final String LINK_FILE = "FILE";

    String address;
    String title;
    HyperlinkType linkType;

    CellStyle linkStyle;

    public WritableHyperlink(String address, String title) {
        this(address, title, LINK_URL);
    }

    public WritableHyperlink(String address, String title, String linkTypeString) {
        this.address = address;
        this.title = title;
        if (LINK_URL.equals(linkTypeString)) {
            linkType = HyperlinkType.URL;
        } else if (LINK_DOCUMENT.equals(linkTypeString)) {
            linkType = HyperlinkType.DOCUMENT;
        } else if (LINK_EMAIL.equals(linkTypeString)) {
            linkType = HyperlinkType.EMAIL;
        } else if (LINK_FILE.equals(linkTypeString)) {
            linkType = HyperlinkType.FILE;
        } else {
            throw new IllegalArgumentException("Link type must be one of the following values: " + LINK_URL + ","
                    + LINK_DOCUMENT + "," + LINK_EMAIL + "," + LINK_FILE);
        }
    }

    public WritableHyperlink(String address, String title, HyperlinkType linkType) {
        this.address = address;
        this.title = title;
        this.linkType = linkType;
    }

    @Override
    public Object writeToCell(Cell cell, Context context) {
        Workbook workbook = cell.getSheet().getWorkbook();
        CreationHelper createHelper = workbook.getCreationHelper();
        Hyperlink hyperlink = createHelper.createHyperlink(linkType);
        hyperlink.setAddress(address);
        cell.setHyperlink(hyperlink);
        cell.setCellValue(title);
        if (linkStyle == null) {
            linkStyle = cell.getCellStyle();
        }
        cell.setCellStyle(linkStyle);
        return cell;
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
