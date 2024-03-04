package org.jxls.builder;

public interface SheetCreater {

    /**
     * Creates a new sheet as a copy of the given source sheet.
     * 
     * @param workbook sheet container
     * @param sourceSheetName name of sheet to be copied
     * @param targetSheetName name of new sheet to be created
     * @return new sheet
     */
    Object createSheet(Object workbook, String sourceSheetName, String targetSheetName);
}
