package com.jxls.writer.transform.poi;

import org.apache.poi.ss.usermodel.Row;

/**
 * @author Leonid Vysochyn
 *         Date: 2/1/12 2:01 PM
 */
public class RowData {
    short height;

    public static RowData createRowData(Row row){
        if( row == null ) return null;
        RowData rowData = new RowData();
        rowData.height = row.getHeight();
        return rowData;
    }

    public short getHeight() {
        return height;
    }

}
