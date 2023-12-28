package org.jxls.formula;

import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.jxls.area.Area;
import org.jxls.common.CellData;
import org.jxls.common.CellRef;
import org.jxls.transform.Transformer;

/**
 * Fast formula processor implementation.
 * It works correctly in 90% of cases and is much more faster than {@link StandardFormulaProcessor}.
 */
public class FastFormulaProcessor extends AbstractFormulaProcessor {

    @Override
    public void processAreaFormulas(Transformer transformer, Area area) {
        transformer.getFormulaCells().forEach(formulaCellData -> {
            if (!(area != null && area.getAreaRef() != null && !area.getAreaRef().getSheetName().equals(formulaCellData.getSheetName()))) {
                processTargetFormulaCells(formulaCellData, transformer, area);
            }
        });
    }

    @Override
    protected void processTargetFormulaCell(int i, CellData formulaCellData, FormulaProcessorContext fpc) {
        processTargetCellRefMap(i, formulaCellData, fpc.targetFormulaCellRef, fpc);
        processJointedCellRefMap(i, fpc);
        processTargetFormula(formulaCellData, fpc);
    }

    private void processTargetCellRefMap(int i, CellData formulaCellData, CellRef targetFormulaCellRef, FormulaProcessorContext fpc) {
        fpc.isFormulaCellRefsEmpty = true;
        // iterate through all the cell references used in the formula
        for (Map.Entry<CellRef, List<CellRef>> cellRefEntry : fpc.targetCellRefMap.entrySet()) {
            // target cells are the cells into which a cell ref from the original formula was transformed
            List<CellRef> targetCells = cellRefEntry.getValue();
            if (targetCells.isEmpty()) {
                continue;
            }
            fpc.isFormulaCellRefsEmpty = false;
            String replacementString;
            // calculate the formula replacement string based on the formula strategy set for the cell
            if (formulaCellData.getFormulaStrategy() == CellData.FormulaStrategy.BY_COLUMN) {
                // BY_COLUMN strategy (non-default) means we will take only cell references in the same column as the original cell
                List<CellRef> targetCellRefs = createTargetCellRefListByColumn(targetFormulaCellRef, targetCells, fpc.usedCellRefs);
                fpc.usedCellRefs.addAll(targetCellRefs);
                replacementString = createTargetCellRef(targetCellRefs);
            } else if (targetCells.size() == fpc.targetFormulaCells.size()) {
                // if the number of the cell reference target cells is the same as the number of cells into which
                // the formula was transformed we assume that a formula target cell should use the
                // corresponding target cell reference
                CellRef targetCellRefCellRef = targetCells.get(i);
                replacementString = targetCellRefCellRef.getCellName();
            } else {
                // trying to group the individual target cell refs used in a formula into a range
                List<List<CellRef>> rangeList = groupByRanges(targetCells, fpc.targetFormulaCells.size());
                if (rangeList.size() == fpc.targetFormulaCells.size()) {
                    // if the number of ranges equals to the number of target formula cells
                    // we assume the formula cells directly map onto ranges and so just taking a corresponding range by index
                    List<CellRef> range = rangeList.get(i);
                    replacementString = createTargetCellRef(range);
                } else {
                    // the range grouping did not succeed and we just use the list of target cells to calculate the replacement string
                    replacementString = createTargetCellRef(targetCells);
                }
            }
            String from = regexJointedLookBehind
                    + sheetNameRegex(cellRefEntry)
                    + getStrictCellNameRegex(Pattern.quote(cellRefEntry.getKey().getCellName()));
            String to = Matcher.quoteReplacement(replacementString);
            fpc.targetFormulaString = fpc.targetFormulaString.replaceAll(from, to);
        }
    }

    private void processJointedCellRefMap(int i, FormulaProcessorContext fpc) {
        fpc.isFormulaJointedCellRefsEmpty = true;
        // iterate through all the jointed cell references used in the formula
        for (Map.Entry<String, List<CellRef>> jointedCellRefEntry : fpc.jointedCellRefMap.entrySet()) {
            List<CellRef> targetCellRefList = jointedCellRefEntry.getValue();
            if (targetCellRefList.isEmpty()) {
                continue;
            }
            fpc.isFormulaJointedCellRefsEmpty = false;
            // trying to group the target cell references into ranges
            List<List<CellRef>> rangeList = groupByRanges(targetCellRefList, fpc.targetFormulaCells.size());
            String replacementString;
            if (rangeList.size() == fpc.targetFormulaCells.size()) {
                // if the number of ranges equals to the number of target formula cells
                // we assume the formula cells directly map onto ranges and so just taking a corresponding range by index
                List<CellRef> range = rangeList.get(i);
                replacementString = createTargetCellRef(range);
            } else {
                replacementString = createTargetCellRef(targetCellRefList);
            }
            fpc.targetFormulaString = fpc.targetFormulaString.replaceAll(Pattern.quote(jointedCellRefEntry.getKey()), replacementString);
        }
    }

    /**
     * @param name -
     * @return regular expression to detect the passed cell name
     */
    private String getStrictCellNameRegex(String name) {
        return "(?<=[^A-Z]|^)" + name + "(?=\\D|$)";
    }
}
