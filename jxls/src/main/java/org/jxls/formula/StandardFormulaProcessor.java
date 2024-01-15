package org.jxls.formula;

import java.util.AbstractMap;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.jxls.area.Area;
import org.jxls.common.AreaRef;
import org.jxls.common.CellData;
import org.jxls.common.CellRef;
import org.jxls.transform.Transformer;

/**
 * This is a standard formula processor implementation which takes into account
 * all the performed cell transformations to properly evaluate all the formulas even
 * for complex templates.
 * <p>However for simple templates you may consider using the {@link FastFormulaProcessor} instead
 * because it is much faster although may not provide the correct results for more complex cases.</p>
 */
public class StandardFormulaProcessor extends AbstractFormulaProcessor {
    private static final int MAX_NUM_ARGS_FOR_SUM = 255;

    /**
     * The method transforms all the formula cells according to the command
     * transformations happened during the area processing
     * @param transformer transformer to use for formula processing
     * @param area - xls area for which the formula processing is invoked
     */
    @Override
    public void processAreaFormulas(Transformer transformer, Area area) {
        transformer.getFormulaCells().forEach(formulaCellData -> {
            if (!(formulaCellData.getArea() == null || !area.getAreaRef().getSheetName().equals(formulaCellData.getSheetName()))) {
                processTargetFormulaCells(formulaCellData, transformer, area);
            }
        });
    }

    @Override
    protected void processTargetFormulaCell(int i, CellData formulaCellData, FormulaProcessorContext fpc) {
        AreaRef formulaSourceAreaRef = formulaCellData.getArea().getAreaRef();
        AreaRef formulaTargetAreaRef = formulaCellData.getTargetParentAreaRef().get(i);
        processTargetCellRefMap(formulaCellData, fpc.targetFormulaCellRef, formulaSourceAreaRef, formulaTargetAreaRef, fpc);
        processJointedCellRefMap(fpc.targetFormulaCellRef, formulaSourceAreaRef, formulaTargetAreaRef, fpc);
        processTargetFormula(formulaCellData, fpc);
    }

    private void processTargetCellRefMap(CellData formulaCellData, CellRef targetFormulaCellRef,
            AreaRef formulaSourceAreaRef, AreaRef formulaTargetAreaRef, FormulaProcessorContext fpc) {
        fpc.isFormulaCellRefsEmpty = true;
        for (Map.Entry<CellRef, List<CellRef>> cellRefEntry : fpc.targetCellRefMap.entrySet()) {
            List<CellRef> targetCells = cellRefEntry.getValue();
            if (targetCells.isEmpty()) {
                continue;
            }
            fpc.isFormulaCellRefsEmpty = false;
            List<CellRef> replacementCells = findFormulaCellRefReplacements(
                    fpc.transformer, targetFormulaCellRef, formulaSourceAreaRef,
                    formulaTargetAreaRef, cellRefEntry);
            if (formulaCellData.getFormulaStrategy() == CellData.FormulaStrategy.BY_COLUMN) {
                // for BY_COLUMN formula strategy we take only a subset of the cells
                replacementCells = createTargetCellRefListByColumn(targetFormulaCellRef, replacementCells, fpc.usedCellRefs);
                fpc.usedCellRefs.addAll(replacementCells);
            }
            String replacementString = createTargetCellRef(replacementCells);
            if (fpc.targetFormulaString.startsWith("SUM")
                    && countOccurences(replacementString, ',') >= MAX_NUM_ARGS_FOR_SUM) {
                // Excel doesn't support more than 255 arguments in functions.
                // Thus, we just concatenate all cells with "+" to have the same effect (see issue B059 for more detail)
                fpc.targetFormulaString = replacementString.replaceAll(",", "+");
            } else {
                String from = regexJointedLookBehind
                        + sheetNameRegex(cellRefEntry)
                        + regexExcludePrefixSymbols
                        + Pattern.quote(cellRefEntry.getKey().getCellName());
                String to = Matcher.quoteReplacement(replacementString);
                fpc.targetFormulaString = fpc.targetFormulaString.replaceAll(from, to);
            }
        }
    }

    private void processJointedCellRefMap(CellRef targetFormulaCellRef, AreaRef formulaSourceAreaRef, AreaRef formulaTargetAreaRef, FormulaProcessorContext fpc) {
        fpc.isFormulaJointedCellRefsEmpty = true;
        // iterate through all the jointed cell references used in the formula
        for (Map.Entry<String, List<CellRef>> jointedCellRefEntry : fpc.jointedCellRefMap.entrySet()) {
            List<CellRef> targetCellRefList = jointedCellRefEntry.getValue();
            if (targetCellRefList.isEmpty()) {
                continue;
            }
            Collections.sort(targetCellRefList);
            fpc.isFormulaJointedCellRefsEmpty = false;
            Map.Entry<CellRef, List<CellRef>> cellRefMapEntryParam =
                    new AbstractMap.SimpleImmutableEntry<>(null, targetCellRefList);
            List<CellRef> replacementCells = findFormulaCellRefReplacements(
                    fpc.transformer, targetFormulaCellRef, formulaSourceAreaRef,
                    formulaTargetAreaRef, cellRefMapEntryParam);
            String replacementString = createTargetCellRef(replacementCells);
            fpc.targetFormulaString = fpc.targetFormulaString.replaceAll(Pattern.quote(jointedCellRefEntry.getKey()), replacementString);
        }
    }

    private List<CellRef> findFormulaCellRefReplacements(Transformer transformer, CellRef targetFormulaCellRef, AreaRef formulaSourceAreaRef,
            AreaRef formulaTargetAreaRef, Map.Entry<CellRef, List<CellRef>> cellReferenceEntry) {
        CellRef cellReference = cellReferenceEntry.getKey();
        List<CellRef> cellReferenceTargets = cellReferenceEntry.getValue();
        if (cellReference != null && !formulaSourceAreaRef.contains(cellReference)) {
            // This cell is outside of the formula cell area. So we just return all the cell
            // reference targets `as is`.
            // However it is possible the cell might be a part of the area replicated in an outer loop
            // In that case we assume we should just take only the target cell ref which belongs to the same area
            // as this ref entry
            CellData cellRefData = transformer.getCellData(cellReference);
            if (cellRefData != null && !cellRefData.getTargetParentAreaRef().isEmpty()) {
                // non-empty means that there was an outer replication of this cell onto new areas
                // we need to find an area which contains both the current formula cell
                // and the cell reference we are searching replacements for
                // since we assume the intention is to use only the target cell reference from the same parent area
                List<CellRef> targetReferences = new ArrayList<>();
                for (AreaRef targetAreaRef : cellRefData.getTargetParentAreaRef()) {
                    if (targetAreaRef.contains(targetFormulaCellRef)) {
                        for (CellRef targetRef : cellReferenceTargets) {
                            if (targetAreaRef.contains(targetRef)) {
                                targetReferences.add(targetRef);
                            }
                        }
                        return targetReferences;
                    }
                }
            }

            return cellReferenceTargets;
        }
        // The source cell reference is inside parent formula area. So let's find target
        // cells related to particular transformation.
        // We'll iterate through all target cell references and find all the ones which
        // belong to the target formula area.
        return findRelevantCellReferences(cellReferenceTargets, formulaTargetAreaRef);
    }

    private List<CellRef> findRelevantCellReferences(List<CellRef> cellReferenceTargets, AreaRef targetFormulaArea) {
        List<CellRef> relevantCellRefs = new ArrayList<>(cellReferenceTargets.size());
        for (CellRef targetCellRef : cellReferenceTargets) {
            if (targetFormulaArea.contains(targetCellRef)) {
                relevantCellRefs.add(targetCellRef);
            }
        }
        return relevantCellRefs;
    }
    
    /**
     * Calculates a number of occurences of a symbol in the string
     * @param string -
     * @param symbol -
     * @return -
     */
    private int countOccurences(String string, char symbol) {
        int count = 0;
        for (int i = 0; i < string.length(); i++) {
            if (string.charAt(i) == symbol) {
                count++;
            }
        }
        return count;
    }
}
