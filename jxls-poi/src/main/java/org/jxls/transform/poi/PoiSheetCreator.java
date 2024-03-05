package org.jxls.transform.poi;

import org.apache.poi.ss.usermodel.HeaderFooter;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jxls.builder.SheetCreator;

public class PoiSheetCreator implements SheetCreator {
    private boolean cloneSheet = true;
    
    /**
     * @return true: use cloneSheet(), false: use createSheet()
     */
    public boolean isCloneSheet() {
        return cloneSheet;
    }

    /**
     * @param cloneSheet true: use cloneSheet(), false: use createSheet()
     */
    public void setCloneSheet(boolean cloneSheet) {
        this.cloneSheet = cloneSheet;
    }

    @Override
    public Object createSheet(Object workbook, String sourceSheetName, String targetSheetName) {
        Sheet src;
        Sheet dest;
        if (cloneSheet && workbook instanceof XSSFWorkbook w) {
            src = w.getSheet(sourceSheetName);
            dest = w.cloneSheet(w.getSheetIndex(sourceSheetName), targetSheetName); // since JXLS 3.0
        } else if (workbook instanceof Workbook w) {
            src = w.getSheet(sourceSheetName);
            dest = w.createSheet(targetSheetName);
        } else {
            throw new IllegalArgumentException("Parameter workbook must be of type Workbook");
        }
        copySheetSetup(src, dest);
        copyPrintSetup(src, dest);
        return dest;
    }

    protected void copySheetSetup(Sheet src, Sheet dest) {
        int srcIndex = src.getWorkbook().getSheetIndex(src);
        String area = src.getWorkbook().getPrintArea(srcIndex);
        if (area != null) {
            int destIndex = dest.getWorkbook().getSheetIndex(dest);
            dest.getWorkbook().setPrintArea(destIndex, area.substring(area.indexOf("!") + 1));
        }
        dest.setAutobreaks(src.getAutobreaks());
        for (int i : src.getColumnBreaks()) {
            dest.setColumnBreak(i);
        }
        for (int i : src.getRowBreaks()) {
            dest.setRowBreak(i);
        }
        dest.setDefaultColumnWidth(src.getDefaultColumnWidth());
        dest.setDefaultRowHeight(src.getDefaultRowHeight());
        dest.setDefaultRowHeightInPoints(src.getDefaultRowHeightInPoints());
        dest.setDisplayGuts(src.getDisplayGuts());
        dest.setFitToPage(src.getFitToPage());
        dest.setForceFormulaRecalculation(src.getForceFormulaRecalculation());
        dest.setHorizontallyCenter(src.getHorizontallyCenter());
        dest.setMargin(Sheet.LeftMargin, src.getMargin(Sheet.LeftMargin));
        dest.setMargin(Sheet.RightMargin, src.getMargin(Sheet.RightMargin));
        dest.setMargin(Sheet.TopMargin, src.getMargin(Sheet.TopMargin));
        dest.setMargin(Sheet.BottomMargin, src.getMargin(Sheet.BottomMargin));
        dest.setPrintGridlines(src.isPrintGridlines());
        dest.setRowSumsBelow(src.getRowSumsBelow());
        dest.setRowSumsRight(src.getRowSumsRight());
        dest.setVerticallyCenter(src.getVerticallyCenter());
        dest.setDisplayFormulas(src.isDisplayFormulas());
        dest.setDisplayGridlines(src.isDisplayGridlines());
        dest.setDisplayZeros(src.isDisplayZeros());
        dest.setPrintGridlines(src.isPrintGridlines());
        dest.setDisplayRowColHeadings(src.isDisplayRowColHeadings());
        dest.setRightToLeft(src.isRightToLeft());
        dest.setRepeatingColumns(src.getRepeatingColumns());
        dest.setRepeatingRows(src.getRepeatingRows());
        copyHeaderOrFooter(src.getHeader(), dest.getHeader());
        copyHeaderOrFooter(src.getFooter(), dest.getFooter());
        if (src instanceof XSSFSheet xs && dest instanceof XSSFSheet xd) {
            xd.getHeaderFooterProperties().setDifferentOddEven(xs.getHeaderFooterProperties().getDifferentOddEven());
            xd.getHeaderFooterProperties().setDifferentFirst(xs.getHeaderFooterProperties().getDifferentFirst());
            xd.getHeaderFooterProperties().setAlignWithMargins(xs.getHeaderFooterProperties().getAlignWithMargins());
            xd.getHeaderFooterProperties().setScaleWithDoc(xs.getHeaderFooterProperties().getScaleWithDoc());
            if (xs.getHeaderFooterProperties().getDifferentOddEven()) {
                copyHeaderOrFooter(xs.getEvenHeader(), xd.getEvenHeader());
                copyHeaderOrFooter(xs.getEvenFooter(), xd.getEvenFooter());
                copyHeaderOrFooter(xs.getOddHeader(), xd.getOddHeader());
                copyHeaderOrFooter(xs.getOddFooter(), xd.getOddFooter());
            }
        }
        dest.setZoom(100);
    }
    
    protected final void copyHeaderOrFooter(HeaderFooter src, HeaderFooter dest) {
        dest.setLeft(src.getLeft());
        dest.setCenter(src.getCenter());
        dest.setRight(src.getRight());
    }
    
    protected void copyPrintSetup(Sheet src, Sheet dest) {
        PrintSetup srcPrintSetup = src.getPrintSetup();
        PrintSetup destPrintSetup = dest.getPrintSetup();
        destPrintSetup.setCopies(srcPrintSetup.getCopies());
        destPrintSetup.setDraft(srcPrintSetup.getDraft());
        destPrintSetup.setFitHeight(srcPrintSetup.getFitHeight());
        destPrintSetup.setFitWidth(srcPrintSetup.getFitWidth());
        destPrintSetup.setFooterMargin(srcPrintSetup.getFooterMargin());
        destPrintSetup.setHeaderMargin(srcPrintSetup.getHeaderMargin());
        destPrintSetup.setHResolution(srcPrintSetup.getHResolution());
        destPrintSetup.setLandscape(srcPrintSetup.getLandscape());
        destPrintSetup.setLeftToRight(srcPrintSetup.getLeftToRight());
        destPrintSetup.setNoColor(srcPrintSetup.getNoColor());
        destPrintSetup.setNoOrientation(srcPrintSetup.getNoOrientation());
        destPrintSetup.setNotes(srcPrintSetup.getNotes());
        destPrintSetup.setPageStart(srcPrintSetup.getPageStart());
        destPrintSetup.setPaperSize(srcPrintSetup.getPaperSize());
        destPrintSetup.setScale(srcPrintSetup.getScale());
        destPrintSetup.setUsePage(srcPrintSetup.getUsePage());
        destPrintSetup.setValidSettings(srcPrintSetup.getValidSettings());
        destPrintSetup.setVResolution(srcPrintSetup.getVResolution());
    }
}
