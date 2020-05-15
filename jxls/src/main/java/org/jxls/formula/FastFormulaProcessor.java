package org.jxls.formula;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.jxls.area.Area;
import org.jxls.common.CellData;
import org.jxls.common.CellRef;
import org.jxls.transform.Transformer;
import org.jxls.util.CellRefUtil;
import org.jxls.util.Util;

/**
 * Fast formula processor implementation.
 * It works correctly in 90% of cases and is much more faster than {@link StandardFormulaProcessor}.
 */
public class FastFormulaProcessor extends AbstractFormulaProcessor {

    // TODO method too long
    // TODO partially similar code to StandardFormulaProcessor
    @Override
    public void processAreaFormulas(Transformer transformer, Area area) {
        Set<CellData> formulaCells = transformer.getFormulaCells();
        for (CellData formulaCellData : formulaCells) {
            if (area != null && area.getAreaRef() != null && !area.getAreaRef().getSheetName().equals(formulaCellData.getSheetName())) {
                continue;
            }
            List<CellRef> targetFormulaCells = formulaCellData.getTargetPos();
            Map<CellRef, List<CellRef>> targetCellRefMap = buildTargetCellRefMap(transformer, area, formulaCellData);
            Map<String, List<CellRef>> jointedCellRefMap = buildJointedCellRefMap(transformer, formulaCellData);
            List<CellRef> usedCellRefs = new ArrayList<>();

            // process all of the result (target) formula cells
            // a result formula cell is a cell into which the original cell with the formula was transformed
            for (int i = 0; i < targetFormulaCells.size(); i++) {
                CellRef targetFormulaCellRef = targetFormulaCells.get(i);
                String targetFormulaString = formulaCellData.getFormula();
                if (formulaCellData.isParameterizedFormulaCell()) {
                    targetFormulaString = formulaCellData.getEvaluatedFormulas().get(i);
                }
                boolean isFormulaCellRefsEmpty = true;
                // iterate through all the cell references used in the formula
                for (Map.Entry<CellRef, List<CellRef>> cellRefEntry : targetCellRefMap.entrySet()) {
                    // target cells are the cells into which a cell ref from the original formula was transformed
                    List<CellRef> targetCells = cellRefEntry.getValue();
                    if (targetCells.isEmpty()) {
                        continue;
                    }
                    isFormulaCellRefsEmpty = false;
                    String replacementString;
                    // calculate the formula replacement string based on the formula strategy set for the cell
                    if (formulaCellData.getFormulaStrategy() == CellData.FormulaStrategy.BY_COLUMN) {
                        // BY_COLUMN strategy (non-default) means we will take only cell references in the same column as the original cell
                        List<CellRef> targetCellRefs = Util.createTargetCellRefListByColumn(targetFormulaCellRef, targetCells, usedCellRefs);
                        usedCellRefs.addAll(targetCellRefs);
                        replacementString = Util.createTargetCellRef(targetCellRefs);
                    } else if (targetCells.size() == targetFormulaCells.size()) {
                        // if the number of the cell reference target cells is the same as the number of cells into which
                        // the formula was transformed we assume that a formula target cell should use the
                        // corresponding target cell reference
                        CellRef targetCellRefCellRef = targetCells.get(i);
                        replacementString = targetCellRefCellRef.getCellName();
                    } else {
                        // trying to group the individual target cell refs used in a formula into a range
                        List<List<CellRef>> rangeList = Util.groupByRanges(targetCells, targetFormulaCells.size());
                        if (rangeList.size() == targetFormulaCells.size()) {
                            // if the number of ranges equals to the number of target formula cells
                            // we assume the formula cells directly map onto ranges and so just taking a corresponding range by index
                            List<CellRef> range = rangeList.get(i);
                            replacementString = Util.createTargetCellRef(range);
                        } else {
                            // the range grouping did not succeed and we just use the list of target cells to calculate the replacement string
                            replacementString = Util.createTargetCellRef(targetCells);
                        }
                    }
                    String from = Util.regexJointedLookBehind
                            + Util.sheetNameRegex(cellRefEntry)
                            + Util.getStrictCellNameRegex(Pattern.quote(cellRefEntry.getKey().getCellName()));
                    String to = Matcher.quoteReplacement(replacementString);
                    targetFormulaString = targetFormulaString.replaceAll(from, to);
                }
                boolean isFormulaJointedCellRefsEmpty = true;
                // iterate through all the jointed cell references used in the formula
                for (Map.Entry<String, List<CellRef>> jointedCellRefEntry : jointedCellRefMap.entrySet()) {
                    List<CellRef> targetCellRefList = jointedCellRefEntry.getValue();
                    if (targetCellRefList.isEmpty()) {
                        continue;
                    }
                    isFormulaJointedCellRefsEmpty = false;
                    // trying to group the target cell references into ranges
                    List<List<CellRef>> rangeList = Util.groupByRanges(targetCellRefList, targetFormulaCells.size());
                    String replacementString;
                    if (rangeList.size() == targetFormulaCells.size()) {
                        // if the number of ranges equals to the number of target formula cells
                        // we assume the formula cells directly map onto ranges and so just taking a corresponding range by index
                        List<CellRef> range = rangeList.get(i);
                        replacementString = Util.createTargetCellRef(range);
                    } else {
                        replacementString = Util.createTargetCellRef(targetCellRefList);
                    }
                    targetFormulaString = targetFormulaString.replaceAll(Pattern.quote(jointedCellRefEntry.getKey()), replacementString);
                }
                String sheetNameReplacementRegex = targetFormulaCellRef.getFormattedSheetName() + CellRefUtil.SHEET_NAME_DELIMITER;
                targetFormulaString = targetFormulaString.replaceAll(sheetNameReplacementRegex, "");
                // if there were no regular or jointed cell references found for this formula use a default value
                // if set or 0
                if (isFormulaCellRefsEmpty && isFormulaJointedCellRefsEmpty
                        && (!formulaCellData.isParameterizedFormulaCell() || formulaCellData.isJointedFormulaCell())) {
                    targetFormulaString = formulaCellData.getDefaultValue() != null ? formulaCellData.getDefaultValue() : "0";
                }
                transformer.setFormula(
                        new CellRef(
                                targetFormulaCellRef.getSheetName(),
                                targetFormulaCellRef.getRow(),
                                targetFormulaCellRef.getCol()),
                        targetFormulaString);
            }
        }
    }
}
