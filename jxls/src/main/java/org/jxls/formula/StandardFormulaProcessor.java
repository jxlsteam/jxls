package org.jxls.formula;

import java.util.AbstractMap;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.jxls.area.Area;
import org.jxls.common.AreaRef;
import org.jxls.common.CellData;
import org.jxls.common.CellRef;
import org.jxls.transform.Transformer;
import org.jxls.util.CellRefUtil;
import org.jxls.util.Util;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * This is a standard formula processor implementation which takes into account
 * all the performed cell transformations to properly evaluate all the formulas even
 * for complex templates.
 * <p>However for simple templates you may consider using the {@link FastFormulaProcessor} instead
 * because it is much faster although may not provide the correct results for more complex cases.</p>
 */
public class StandardFormulaProcessor extends AbstractFormulaProcessor {
    private static Logger logger = LoggerFactory.getLogger(StandardFormulaProcessor.class);
    private static final int MAX_NUM_ARGS_FOR_SUM = 255;

    // TODO method too long
    /**
     * The method transforms all the formula cells according to the command
     * transformations happened during the area processing
     * @param transformer transformer to use for formula processing
     * @param area - xls area for which the formula processing is invoked
     */
    @Override
    public void processAreaFormulas(Transformer transformer, Area area) {
        Set<CellData> formulaCells = transformer.getFormulaCells();
        for (CellData formulaCellData : formulaCells) {
            logger.debug("Processing formula cell {}", formulaCellData);
            List<CellRef> targetFormulaCells = formulaCellData.getTargetPos();
            Map<CellRef, List<CellRef>> targetCellRefMap = buildTargetCellRefMap(transformer, area, formulaCellData);
            Map<String, List<CellRef>> jointedCellRefMap = buildJointedCellRefMap(transformer, formulaCellData);
            List<CellRef> usedCellRefs = new ArrayList<>();

            // process all of the result (target) formula cells
            // a result formula cell is a cell into which the original cell with the formula was transformed
            for (int i = 0; i < targetFormulaCells.size(); i++) {
                CellRef targetFormulaCellRef = targetFormulaCells.get(i);
                String targetFormulaString = formulaCellData.getFormula();
                if (formulaCellData.getArea() == null) {
                    continue;
                }
                AreaRef formulaSourceAreaRef = formulaCellData.getArea().getAreaRef();
                AreaRef formulaTargetAreaRef = formulaCellData.getTargetParentAreaRef().get(i);
                boolean isFormulaCellRefsEmpty = true;
                for (Map.Entry<CellRef, List<CellRef>> cellRefEntry : targetCellRefMap.entrySet()) {
                    List<CellRef> targetCells = cellRefEntry.getValue();
                    if (targetCells.isEmpty()) {
                        continue;
                    }
                    isFormulaCellRefsEmpty = false;
                    List<CellRef> replacementCells = findFormulaCellRefReplacements(formulaSourceAreaRef,
                            formulaTargetAreaRef, cellRefEntry);
                    if (formulaCellData.getFormulaStrategy() == CellData.FormulaStrategy.BY_COLUMN) {
                        // for BY_COLUMN formula strategy we take only a subset of the cells
                        replacementCells = Util.createTargetCellRefListByColumn(targetFormulaCellRef, replacementCells,
                                usedCellRefs);
                        usedCellRefs.addAll(replacementCells);
                    }
                    String replacementString = Util.createTargetCellRef(replacementCells);
                    if (targetFormulaString.startsWith("SUM")
                            && Util.countOccurences(replacementString, ',') >= MAX_NUM_ARGS_FOR_SUM) {
                        // Excel doesn't support more than 255 arguments in functions.
                        // Thus, we just concatenate all cells with "+" to have the same effect (see issue#59 for more detail)
                        targetFormulaString = replacementString.replaceAll(",", "+");
                    } else {
                        String from = Util.regexJointedLookBehind + Util.sheetNameRegex(cellRefEntry)
                                + Util.regexExcludePrefixSymbols
                                + Pattern.quote(cellRefEntry.getKey().getCellName());
                        String to = Matcher.quoteReplacement(replacementString);
                        targetFormulaString = targetFormulaString.replaceAll(from, to);
                    }
                }
                boolean isFormulaJointedCellRefsEmpty = true;
                // iterate through all the jointed cell references used in the formula
                for (Map.Entry<String, List<CellRef>> jointedCellRefEntry : jointedCellRefMap.entrySet()) {
                    List<CellRef> targetCellRefList = jointedCellRefEntry.getValue();
                    Collections.sort(targetCellRefList);
                    if (targetCellRefList.isEmpty()) {
                        continue;
                    }
                    isFormulaJointedCellRefsEmpty = false;
                    Map.Entry<CellRef, List<CellRef>> cellRefMapEntryParam =
                            new AbstractMap.SimpleImmutableEntry<CellRef, List<CellRef>>(null, targetCellRefList);
                    List<CellRef> replacementCells = findFormulaCellRefReplacements(formulaSourceAreaRef,
                            formulaTargetAreaRef, cellRefMapEntryParam);
                    String replacementString = Util.createTargetCellRef(replacementCells);
                    targetFormulaString = targetFormulaString.replaceAll(Pattern.quote(jointedCellRefEntry.getKey()), replacementString);
                }
                String sheetNameReplacementRegex = Pattern.quote(targetFormulaCellRef.getFormattedSheetName() + CellRefUtil.SHEET_NAME_DELIMITER);
                targetFormulaString = targetFormulaString.replaceAll(sheetNameReplacementRegex, "");
                // if there were no regular or jointed cell references found for this formula use a default value
                // if set or 0
                if (isFormulaCellRefsEmpty && isFormulaJointedCellRefsEmpty) {
                    targetFormulaString = formulaCellData.getDefaultValue() != null ? formulaCellData.getDefaultValue() : "0";
                }
                if (!targetFormulaString.isEmpty()) {
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

    private List<CellRef> findFormulaCellRefReplacements(AreaRef formulaSourceAreaRef, AreaRef formulaTargetAreaRef,
            Map.Entry<CellRef, List<CellRef>> cellReferenceEntry) {
        CellRef cellReference = cellReferenceEntry.getKey();
        List<CellRef> cellReferenceTargets = cellReferenceEntry.getValue();
        if (cellReference != null && !formulaSourceAreaRef.contains(cellReference)) {
            // This cell is outside of the formula cell area. So we just return all the cell
            // reference targets `as is`.
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
}
