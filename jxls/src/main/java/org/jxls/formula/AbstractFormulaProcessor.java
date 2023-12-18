package org.jxls.formula;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.jxls.area.Area;
import org.jxls.common.CellData;
import org.jxls.common.CellRef;
import org.jxls.transform.Transformer;

/**
 * Partial implementation of {@link FormulaProcessor} interface
 * It implements only some helper methods to allow their reuse for
 * {@link FastFormulaProcessor} and {@link StandardFormulaProcessor}
 */
public abstract class AbstractFormulaProcessor implements FormulaProcessor {

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
            List<CellRef> jointedCellRefList = new ArrayList<CellRef>();
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
    
    public static final String regexJointedLookBehind = "(?<!U_\\([^)]{0,100})";
    public static final String regexSimpleCellRef = "[a-zA-Z]+[0-9]+";
    public static final String regexCellRef = "([a-zA-Z_]+[a-zA-Z0-9_]*![a-zA-Z]+[0-9]+|(?<!\\d)[a-zA-Z]+[0-9]+|'[^?\\\\/:'*]+'![a-zA-Z]+[0-9]+)";
    public static final String regexCellRefExcludingJointed = regexJointedLookBehind + regexCellRef;
    private static final Pattern regexCellRefExcludingJointedPattern = Pattern.compile(regexCellRefExcludingJointed);
    private static final Pattern regexCellRefPattern = Pattern.compile(regexCellRef);
    public static final String regexJointedCellRef = "U_\\([^\\)]+\\)";
    public static final Pattern regexJointedCellRefPattern = Pattern.compile(regexJointedCellRef);
    public static final String regexExcludePrefixSymbols = "(?<!\\w)";

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
        List<String> cellRefs = new ArrayList<String>();
        if (str != null) {
            Matcher cellRefMatcher = pattern.matcher(str);
            while (cellRefMatcher.find()) {
                cellRefs.add(cellRefMatcher.group());
            }
        }
        return cellRefs;
    }
}
