package org.jxls;

import java.io.File;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.ConditionalFormatting;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 * Class that encapsulates POI for testing Excel file contents.
 */
public class TestWorkbook implements AutoCloseable {
    private Workbook workbook;
    private Sheet sheet;
    
    public TestWorkbook(File file) {
        try {
            workbook = WorkbookFactory.create(file, null, true);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * Following operations operate on the given sheet
     * @param name exact visible name
     */
    public void selectSheet(String name) {
		sheet = workbook.getSheet(name);
		if (sheet == null) {
			throw new IllegalArgumentException("Sheet \"" + name + "\" does not exist!");
		} else if (sheet instanceof XSSFSheet x
				&& !SheetVisibility.VISIBLE.equals(workbook.getSheetVisibility(workbook.getSheetIndex(name)))) {
			throw new IllegalArgumentException("Sheet \"" + name + "\" is not visible!");
		}
    }

    /**
     * Following operations operate on the given sheet
     * @param index starts with 0
     */
    public void selectSheet(int index) {
        sheet = workbook.getSheetAt(index);
    }

    /**
     * Expects text cell and returns its content.
     * @param row starts with 1
     * @param column 1 = A
     * @return String
     */
    public String getCellValueAsString(int row, int column) {
        try {
            return sheet.getRow(row - 1).getCell(column - 1).getStringCellValue();
        } catch (NullPointerException e) {
            return null;
        }
    }

    /**
     * Expects (possibly formatted) text cell and returns its content.
     * @param row starts with 1
     * @param column 1 = A
     * @return RichTextString.toString()
     */
    public String getCellValueAsRichString(int row, int column) {
        try {
            return sheet.getRow(row - 1).getCell(column - 1).getRichStringCellValue().toString();
        } catch (NullPointerException e) {
            return null;
        }
    }
    
    /**
     * Expects (possibly formatted) text cell and returns its numFormattingRuns value.
     * @param row starts with 1
     * @param column 1 = A
     * @return numFormattingRuns value
     */
    public int getCellValueAsRichStringNumFormattingRuns(int row, int column) {
        return sheet.getRow(row - 1).getCell(column - 1).getRichStringCellValue().numFormattingRuns();
    }

    /**
     * Expects formula cell and returns the formula string.
     * @param row starts with 1
     * @param column 1 = A
     * @return String
     */
    public String getFormulaString(int row, int column) {
        try {
            return sheet.getRow(row - 1).getCell(column - 1).getCellFormula();
        } catch (NullPointerException e) {
            return null;
        }
    }

    /**
     * Expects numeric cell and returns its double value.
     * @param row starts with 1
     * @param column 1 = A
     * @return double
     */
    public double getCellValueAsDouble(int row, int column) {
        return Double.valueOf(sheet.getRow(row - 1).getCell(column - 1).getNumericCellValue()).doubleValue();
    }

    /**
     * Expects date cell and returns its value as {@link LocalDateTime}.
     * @param row starts with 1
     * @param column 1 = A
     * @return LocalDateTime
     */
    public LocalDateTime getCellValueAsLocalDateTime(int row, int column) {
        try {
            return sheet.getRow(row - 1).getCell(column - 1).getLocalDateTimeCellValue();
        } catch (NullPointerException e) {
            return null;
        }
    }

    /**
     * @param row starts with 1
     * @return row height in Twips
     */
    public short getRowHeight(int row) {
        return sheet.getRow(row - 1).getHeight();
    }
    
    public int getColumnWidth(int column) {
    	return sheet.getColumnWidth(column - 1);
    }

    public int getConditionalFormattingSize() {
        int conditionalFormattingCount = 0;
        SheetConditionalFormatting sheetConditionalFormatting = sheet.getSheetConditionalFormatting();
        for (int i = 0; i < sheetConditionalFormatting.getNumConditionalFormattings(); i++) {
            ConditionalFormatting conditionalFormatting = sheetConditionalFormatting.getConditionalFormattingAt(i);
            CellRangeAddress[] ranges = conditionalFormatting.getFormattingRanges();
            if (ranges.length > 0) {
                conditionalFormattingCount++;
            }
        }
        return conditionalFormattingCount;
    }
    
    public List<String> getConditionalFormattingRanges() {
        List<String> ret = new ArrayList<>();
        SheetConditionalFormatting sheetConditionalFormatting = sheet.getSheetConditionalFormatting();
        for (int i = 0; i < sheetConditionalFormatting.getNumConditionalFormattings(); i++) {
            ConditionalFormatting conditionalFormatting = sheetConditionalFormatting.getConditionalFormattingAt(i);
            CellRangeAddress[] ranges = conditionalFormatting.getFormattingRanges();
            for (CellRangeAddress c : ranges) {
                ret.add(c.formatAsString());
            }
        }
        return ret;
    }
    
    public boolean isForceFormulaRecalculation() {
        return workbook.getForceFormulaRecalculation();
    }
    
    @Override
    public void close() {
        if (workbook != null) {
            try {
                workbook.close();
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }
    }
}
