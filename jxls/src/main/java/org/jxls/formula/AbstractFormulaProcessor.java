package org.jxls.formula;

import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.jxls.area.Area;
import org.jxls.common.CellData;
import org.jxls.common.CellRef;
import org.jxls.common.CellRefColPrecedenceComparator;
import org.jxls.common.CellRefRowPrecedenceComparator;
import org.jxls.transform.Transformer;
import org.jxls.util.CellRefUtil;

/**
 * Partial implementation of {@link FormulaProcessor} interface
 * It implements only some helper methods to allow their reuse for
 * {@link FastFormulaProcessor} and {@link StandardFormulaProcessor}
 */
public abstract class AbstractFormulaProcessor implements FormulaProcessor {
    protected static final String regexJointedLookBehind = "(?<!U_\\([^)]{0,100})";
    public static final String regexCellRef = "([a-zA-Z_]+[a-zA-Z0-9_]*![a-zA-Z]+[0-9]+|(?<!\\d)[a-zA-Z]+[0-9]+|'[^?\\\\/:'*]+'![a-zA-Z]+[0-9]+)";
    private static final String regexCellRefExcludingJointed = regexJointedLookBehind + regexCellRef;
    private static final Pattern regexCellRefExcludingJointedPattern = Pattern.compile(regexCellRefExcludingJointed);
    private static final Pattern regexCellRefPattern = Pattern.compile(regexCellRef);
    private static final String regexJointedCellRef = "U_\\([^\\)]+\\)";
    public static final Pattern regexJointedCellRefPattern = Pattern.compile(regexJointedCellRef);
    protected static final String regexExcludePrefixSymbols = "(?<!\\w)";

    // building a map of all the cell references used in a formula
    // and the result cells into which they were transformed during the workbook processing
    protected Map<CellRef, List<CellRef>> buildTargetCellRefMap(Transformer transformer, Area area, CellData formulaCellData) {
        Map<CellRef, List<CellRef>> targetCellRefMap = new LinkedHashMap<>();
        // getting a list of cell names used in the formula
        List<String> formulaCellRefs = getFormulaCellRefs(formulaCellData.getFormula());
        // for each cell ref build a list of cells it was transformed into
        for (String cellRef : formulaCellRefs) {
            CellRef pos = new CellRef(cellRef);
            if (pos.isValid()) {
                if (pos.getSheetName() == null) {
                    pos.setSheetName(formulaCellData.getSheetName());
                    pos.setIgnoreSheetNameInFormat(true);
                }
                List<CellRef> targetCellDataList = transformer.getTargetCellRef(pos);
                // if the cell was not transformed into any cell and it is outside of the source area of the current sheet
                // we set the target cell to be the same as the source cell
                if (targetCellDataList.isEmpty() &&
                        area != null &&
                        !area.getAreaRef().contains(pos)) {
                    targetCellDataList.add(pos);
                }
                targetCellRefMap.put(pos, targetCellDataList);
            }
        }
        return targetCellRefMap;
    }

    // building a map of all the "jointed cell" references used in a formula
    // see Util.getJointedCellRefs() method for an explanation of what a jointed cell ref is
    protected Map<String, List<CellRef>> buildJointedCellRefMap(Transformer transformer, CellData formulaCellData) {
        Map<String, List<CellRef>> jointedCellRefMap = new LinkedHashMap<>();
        // getting a list of the jointed cell references used in the formula
        List<String> jointedCellRefs = getJointedCellRefs(formulaCellData.getFormula());
        // for each jointed cell ref build a list of cells into which individual cell names were transformed to
        for (String jointedCellRef : jointedCellRefs) {
            List<String> nestedCellRefs = getCellRefsFromJointedCellRef(jointedCellRef);
            List<CellRef> jointedCellRefList = new ArrayList<>();
            for (String cellRef : nestedCellRefs) {
                CellRef pos = new CellRef(cellRef);
                if (pos.getSheetName() == null) {
                    pos.setSheetName(formulaCellData.getSheetName());
                    pos.setIgnoreSheetNameInFormat(true);
                }
                List<CellRef> targetCellDataList = transformer.getTargetCellRef(pos);
                jointedCellRefList.addAll(targetCellDataList);
            }
            jointedCellRefMap.put(jointedCellRef, jointedCellRefList);
        }
        return jointedCellRefMap;
    }

    /**
     * Parses a "jointed cell" reference and extracts individual cell references
     * @param jointedCellRef a jointed cell reference to parse
     * @return a list of cell names extracted from the jointed cell reference
     */
    protected List<String> getCellRefsFromJointedCellRef(String jointedCellRef) {
        return getStringPartsByPattern(jointedCellRef, regexCellRefPattern);
    }
    
    /**
     * Parses a formula and returns a list of cell names used in it
     * E.g. for formula "B4*(1+C4)" the returned list will contain "B4", "C4"
     * @param formula string
     * @return a list of cell names used in the formula
     */
    public static List<String> getFormulaCellRefs(String formula) {
        return getStringPartsByPattern(formula, regexCellRefExcludingJointedPattern);
    }

    /**
     * Parses a formula to extract a list of so called "jointed cells"
     * The jointed cells are cells combined with a special notation "U_(cell1, cell2)" into a single cell
     * They are used in formulas like this "$[SUM(U_(F8,F13))]".
     * Here the formula will use both F8 and F13 source cells to calculate the sum
     * @param formula a formula string to parse
     * @return a list of jointed cells used in the formula
     */
    protected List<String> getJointedCellRefs(String formula) {
        return getStringPartsByPattern(formula, regexJointedCellRefPattern);
    }

    private static List<String> getStringPartsByPattern(String str, Pattern pattern) {
        List<String> cellRefs = new ArrayList<>();
        if (str != null) {
            Matcher cellRefMatcher = pattern.matcher(str);
            while (cellRefMatcher.find()) {
                cellRefs.add(cellRefMatcher.group());
            }
        }
        return cellRefs;
    }

    /**
     * Combines a list of cell references into a range
     * E.g. for cell references A1, A2, A3, A4 it returns A1:A4
     * @param targetCellDataList -
     * @return a range containing all the cell references if such range exists or otherwise the passed cells separated by commas
     */
    static String createTargetCellRef(List<CellRef> targetCellDataList) {
        // testcase: UtilCreateTargetCellRefTest. Can be optimized in Java 8.
        if (targetCellDataList == null) {
            return "";
        }
        int size = targetCellDataList.size();
        if (size == 0) {
            return "";
        } else if (size == 1) {
            return targetCellDataList.get(0).getCellName();
        }
        
        // falsify if same sheet
        for (int i = 0; i < size - 1; i++) {
            if (!targetCellDataList.get(i).getSheetName().equals(targetCellDataList.get(i + 1).getSheetName())) {
                return buildCellRefsString(targetCellDataList);
            }
        }
        
        // falsify if rectangular
        CellRef upperLeft = targetCellDataList.get(0);
        CellRef lowerRight = targetCellDataList.get(size - 1);
        int rowCount = lowerRight.getRow() - upperLeft.getRow() + 1;
        int colCount = lowerRight.getCol() - upperLeft.getCol() + 1;
        if (size != colCount * rowCount) {
            return buildCellRefsString(targetCellDataList);
        }
        // Fast exit if horizontal or vertical
        if (rowCount == 1 || colCount == 1) {
            return upperLeft.getCellName() + ":" + lowerRight.getCellName();
        }
        
        // Hole in rectangle with same cell count check
        // Check if upperLeft is most upper cell and most left cell. And check if lowerRight is most lower cell and most right cell.
        int minRow = upperLeft.getRow();
        int minCol = upperLeft.getCol();
        int maxRow = minRow;
        int maxCol = minCol;
        for (CellRef cell : targetCellDataList) {
            if (cell.getCol() < minCol) {
                minCol = cell.getCol();
            }
            if (cell.getCol() > maxCol) {
                maxCol = cell.getCol();
            }
            if (cell.getRow() < minRow) {
                minRow = cell.getRow();
            }
            if (cell.getRow() > maxRow) {
                maxRow = cell.getRow();
            }
        }
        if (!(maxRow == lowerRight.getRow() && minRow == upperLeft.getRow() && maxCol == lowerRight.getCol() && minCol == upperLeft.getCol())) {
            return buildCellRefsString(targetCellDataList);
        }

        // Selection is either vertical, horizontal line or rectangular -> same return structure in each case
        return upperLeft.getCellName() + ":" + lowerRight.getCellName();
    }

    private static String buildCellRefsString(List<CellRef> cellRefs) {
        String reply = "";
        for (CellRef cellRef : cellRefs) {
            reply += "," + cellRef.getCellName();
        }
        return reply.substring(1);
    }
    
    /**
     * Groups a list of cell references into a list ranges which can be used in a formula substitution
     * @param cellRefList a list of cell references
     * @param targetRangeCount a number of ranges to use when grouping
     * @return a list of cell ranges grouped by row or by column
     */
    protected List<List<CellRef>> groupByRanges(List<CellRef> cellRefList, int targetRangeCount) {
        List<List<CellRef>> colRanges = groupByColRange(cellRefList);
        if (targetRangeCount == 0 || colRanges.size() == targetRangeCount) {
            return colRanges;
        }
        List<List<CellRef>> rowRanges = groupByRowRange(cellRefList);
        if (rowRanges.size() == targetRangeCount) {
            return rowRanges;
        } else {
            return colRanges;
        }
    }

    /**
     * Groups a list of cell references in a column into a list of ranges
     * @param cellRefList -
     * @return a list of cell reference groups
     */
    protected List<List<CellRef>> groupByColRange(List<CellRef> cellRefList) {
        List<List<CellRef>> rangeList = new ArrayList<>();
        if (cellRefList == null || cellRefList.size() == 0) {
            return rangeList;
        }
        List<CellRef> cellRefListCopy = new ArrayList<>(cellRefList);
        Collections.sort(cellRefListCopy, new CellRefColPrecedenceComparator());

        String sheetName = cellRefListCopy.get(0).getSheetName();
        int row = cellRefListCopy.get(0).getRow();
        int col = cellRefListCopy.get(0).getCol();
        List<CellRef> currentRange = new ArrayList<>();
        currentRange.add(cellRefListCopy.get(0));
        boolean rangeComplete = false;
        for (int i = 1; i < cellRefListCopy.size(); i++) {
            CellRef cellRef = cellRefListCopy.get(i);
            if (!cellRef.getSheetName().equals(sheetName)) {
                rangeComplete = true;
            } else {
                int rowDelta = cellRef.getRow() - row;
                int colDelta = cellRef.getCol() - col;
                if (rowDelta == 1 && colDelta == 0) {
                    currentRange.add(cellRef);
                } else {
                    rangeComplete = true;
                }
            }
            sheetName = cellRef.getSheetName();
            row = cellRef.getRow();
            col = cellRef.getCol();
            if (rangeComplete) {
                rangeList.add(currentRange);
                currentRange = new ArrayList<>();
                currentRange.add(cellRef);
                rangeComplete = false;
            }
        }
        rangeList.add(currentRange);
        return rangeList;
    }

    /**
     * Groups a list of cell references in a row into a list of ranges
     * @param cellRefList -
     * @return -
     */
    protected List<List<CellRef>> groupByRowRange(List<CellRef> cellRefList) {
        List<List<CellRef>> rangeList = new ArrayList<>();
        if (cellRefList == null || cellRefList.size() == 0) {
            return rangeList;
        }
        List<CellRef> cellRefListCopy = new ArrayList<>(cellRefList);
        Collections.sort(cellRefListCopy, new CellRefRowPrecedenceComparator());

        String sheetName = cellRefListCopy.get(0).getSheetName();
        int row = cellRefListCopy.get(0).getRow();
        int col = cellRefListCopy.get(0).getCol();
        List<CellRef> currentRange = new ArrayList<>();
        currentRange.add(cellRefListCopy.get(0));
        boolean rangeComplete = false;
        for (int i = 1; i < cellRefListCopy.size(); i++) {
            CellRef cellRef = cellRefListCopy.get(i);
            if (!cellRef.getSheetName().equals(sheetName)) {
                rangeComplete = true;
            } else {
                int rowDelta = cellRef.getRow() - row;
                int colDelta = cellRef.getCol() - col;
                if (colDelta == 1 && rowDelta == 0) {
                    currentRange.add(cellRef);
                } else {
                    rangeComplete = true;
                }
            }
            sheetName = cellRef.getSheetName();
            row = cellRef.getRow();
            col = cellRef.getCol();
            if (rangeComplete) {
                rangeList.add(currentRange);
                currentRange = new ArrayList<>();
                currentRange.add(cellRef);
                rangeComplete = false;
            }
        }
        rangeList.add(currentRange);
        return rangeList;
    }

    /**
     * @param cellRefEntry -
     * @return the sheet name regular expression string
     */
    protected String sheetNameRegex(Map.Entry<CellRef, List<CellRef>> cellRefEntry) {
        return (cellRefEntry.getKey().isIgnoreSheetNameInFormat() ? "(?<!!)" : "");
    }
    
    /**
     * Creates a list of target formula cell references
     * @param targetFormulaCellRef -
     * @param targetCells -
     * @param cellRefsToExclude -
     * @return -
     */
    protected List<CellRef> createTargetCellRefListByColumn(CellRef targetFormulaCellRef, List<CellRef> targetCells, List<CellRef> cellRefsToExclude) {
        List<CellRef> resultCellList = new ArrayList<>();
        int col = targetFormulaCellRef.getCol();
        for (CellRef targetCell : targetCells) {
            if (targetCell.getCol() == col
                    && targetCell.getRow() < targetFormulaCellRef.getRow()
                    && !cellRefsToExclude.contains(targetCell)) {
                resultCellList.add(targetCell);
            }
        }
        return resultCellList;
    }
    
    static class FormulaProcessorContext {
        Transformer transformer;
        List<CellRef> targetFormulaCells;
        Map<CellRef, List<CellRef>> targetCellRefMap;
        Map<String, List<CellRef>> jointedCellRefMap;
        List<CellRef> usedCellRefs = new ArrayList<>();
        boolean isFormulaCellRefsEmpty;
        boolean isFormulaJointedCellRefsEmpty;
        String targetFormulaString;
        CellRef targetFormulaCellRef;
    }

    protected void processTargetFormulaCells(CellData formulaCellData, Transformer transformer, Area area) {
        transformer.getLogger().debug("Processing formula cell " + formulaCellData);
        FormulaProcessorContext fpc = createFormulaProcessorContext(formulaCellData, transformer, area);

        // process all of the result (target) formula cells
        // a result formula cell is a cell into which the original cell with the formula was transformed
        for (int i = 0; i < fpc.targetFormulaCells.size(); i++) {
            fpc.targetFormulaCellRef = fpc.targetFormulaCells.get(i);
            fpc.targetFormulaString = formulaCellData.getFormula();
            if (formulaCellData.isParameterizedFormulaCell() && i < formulaCellData.getEvaluatedFormulas().size()) {
                fpc.targetFormulaString = formulaCellData.getEvaluatedFormulas().get(i);
            }
            processTargetFormulaCell(i, formulaCellData, fpc);
        }
    }
    
    protected abstract void processTargetFormulaCell(int i, CellData formulaCellData, FormulaProcessorContext fpc);
    
    protected FormulaProcessorContext createFormulaProcessorContext(CellData formulaCellData, Transformer transformer, Area area) {
        FormulaProcessorContext fpc = new FormulaProcessorContext();
        fpc.transformer = transformer;
        fpc.targetFormulaCells = formulaCellData.getTargetPos();
        fpc.targetCellRefMap = buildTargetCellRefMap(transformer, area, formulaCellData);
        fpc.jointedCellRefMap = buildJointedCellRefMap(transformer, formulaCellData);
        return fpc;
    }

    protected void processTargetFormula(CellData formulaCellData, FormulaProcessorContext fpc) {
        String sheetNameReplacementRegex = Pattern.quote(fpc.targetFormulaCellRef.getFormattedSheetName() + CellRefUtil.SHEET_NAME_DELIMITER);
        fpc.targetFormulaString = fpc.targetFormulaString.replaceAll(sheetNameReplacementRegex, "");
        // if there were no regular or jointed cell references found for this formula use a default value
        // if set or 0
        if (fpc.isFormulaCellRefsEmpty && fpc.isFormulaJointedCellRefsEmpty
                && (!formulaCellData.isParameterizedFormulaCell() || formulaCellData.isJointedFormulaCell())) {
            fpc.targetFormulaString = formulaCellData.getDefaultValue() != null ? formulaCellData.getDefaultValue() : "0";
        }
        if (!fpc.targetFormulaString.isEmpty()) {
            fpc.transformer.setFormula(new CellRef(fpc.targetFormulaCellRef), fpc.targetFormulaString);
        }
    }
}
