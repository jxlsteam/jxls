package org.jxls.formula;

import org.jxls.common.AreaRef;
import org.jxls.common.CellData;
import org.jxls.common.CellRef;
import org.jxls.transform.Transformer;
import org.jxls.util.CellRefUtil;
import org.jxls.util.Util;

import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * This is a standard formula processor implementation which takes into account all performed cell
 * transformations to properly evaluate all the formulas even for complex templates.
 * But it is very-very slow.
 * In many cases it is better to use {@link FastFormulaProcessor} as it is much-much faster
 * although may produce incorrect results in some specific cases
 */
public class StandardFormulaProcessor implements FormulaProcessor {
    @Override
    public void processAreaFormulas(Transformer transformer) {
        Set<CellData> formulaCells = transformer.getFormulaCells();
        for (CellData formulaCellData : formulaCells) {
            List<String> formulaCellRefs = Util.getFormulaCellRefs(formulaCellData.getFormula());
            List<String> jointedCellRefs = Util.getJointedCellRefs(formulaCellData.getFormula());
            List<CellRef> targetFormulaCells = formulaCellData.getTargetPos();
            Map<CellRef, List<CellRef>> targetCellRefMap = new LinkedHashMap<>();
            Map<String, List<CellRef>> jointedCellRefMap = new LinkedHashMap<>();
            for (String cellRef : formulaCellRefs) {
                CellRef pos = new CellRef(cellRef);
                if( pos.isValid() ) {
                    if (pos.getSheetName() == null) {
                        pos.setSheetName(formulaCellData.getSheetName());
                        pos.setIgnoreSheetNameInFormat(true);
                    }
                    List<CellRef> targetCellDataList = transformer.getTargetCellRef(pos);
                    targetCellRefMap.put(pos, targetCellDataList);
                }
            }
            for (String jointedCellRef : jointedCellRefs) {
                List<String> nestedCellRefs = Util.getCellRefsFromJointedCellRef(jointedCellRef);
                List<CellRef> jointedCellRefList = new ArrayList<CellRef>();
                for (String cellRef : nestedCellRefs) {
                    CellRef pos = new CellRef(cellRef);
                    if(pos.getSheetName() == null ){
                        pos.setSheetName(formulaCellData.getSheetName());
                        pos.setIgnoreSheetNameInFormat(true);
                    }
                    List<CellRef> targetCellDataList = transformer.getTargetCellRef(pos);

                    jointedCellRefList.addAll(targetCellDataList);
                }
                jointedCellRefMap.put(jointedCellRef, jointedCellRefList);
            }
            List<CellRef> usedCellRefs = new ArrayList<>();
            for (int i = 0; i < targetFormulaCells.size(); i++) {
                CellRef targetFormulaCellRef = targetFormulaCells.get(i);
                String targetFormulaString = formulaCellData.getFormula();
                AreaRef formulaSourceAreaRef = formulaCellData.getArea().getAreaRef();
                AreaRef formulaTargetAreaRef = formulaCellData.getTargetParentAreaRef().get(i);
                boolean isFormulaCellRefsEmpty = true;
                for (Map.Entry<CellRef, List<CellRef>> cellRefEntry : targetCellRefMap.entrySet()) {
                    List<CellRef> targetCells = cellRefEntry.getValue();
                    if( targetCells.isEmpty() ) {
                        continue;
                    }
                    isFormulaCellRefsEmpty = false;
                    List<CellRef> replacementCells = findFormulaCellRefReplacements(formulaSourceAreaRef, formulaTargetAreaRef, cellRefEntry);
                    if( formulaCellData.getFormulaStrategy() == CellData.FormulaStrategy.BY_COLUMN ){
                        replacementCells = Util.createTargetCellRefListByColumn(targetFormulaCellRef, replacementCells, usedCellRefs);
                        usedCellRefs.addAll(replacementCells);
                    }
                    String replacementString = Util.createTargetCellRef(replacementCells);
                    targetFormulaString = targetFormulaString.replaceAll(Util.regexJointedLookBehind + Util.sheetNameRegex(cellRefEntry) + Pattern.quote(cellRefEntry.getKey().getCellName()), Matcher.quoteReplacement(replacementString));
                }
                for (Map.Entry<String, List<CellRef>> jointedCellRefEntry : jointedCellRefMap.entrySet()) {
                    List<CellRef> targetCellRefList = jointedCellRefEntry.getValue();
                    Collections.sort(targetCellRefList);
                    if( targetCellRefList.isEmpty() ) {
                        continue;
                    }
                    isFormulaCellRefsEmpty = false;
                    Map.Entry<CellRef, List<CellRef>> cellRefMapEntryParam = new AbstractMap.SimpleImmutableEntry<CellRef, List<CellRef>>(null, targetCellRefList);
                    List<CellRef> replacementCells = findFormulaCellRefReplacements(formulaSourceAreaRef, formulaTargetAreaRef, cellRefMapEntryParam);
                    String replacementString = Util.createTargetCellRef(replacementCells);
                    targetFormulaString = targetFormulaString.replaceAll(Pattern.quote(jointedCellRefEntry.getKey()), replacementString);
                }
                String sheetNameReplacementRegex = targetFormulaCellRef.getFormattedSheetName() + CellRefUtil.SHEET_NAME_DELIMITER;
                targetFormulaString = targetFormulaString.replaceAll(sheetNameReplacementRegex, "");
                if( isFormulaCellRefsEmpty ){
                    targetFormulaString = formulaCellData.getDefaultValue() != null ? formulaCellData.getDefaultValue() : "0";
                }
                transformer.setFormula(new CellRef(targetFormulaCellRef.getSheetName(), targetFormulaCellRef.getRow(), targetFormulaCellRef.getCol()), targetFormulaString);
            }
        }
    }

    private List<CellRef> findFormulaCellRefReplacements(AreaRef formulaSourceAreaRef, AreaRef formulaTargetAreaRef, Map.Entry<CellRef, List<CellRef>> cellReferenceEntry) {
        CellRef cellReference = cellReferenceEntry.getKey();
        List<CellRef> cellReferenceTargets = cellReferenceEntry.getValue();
        if( cellReference != null && !formulaSourceAreaRef.contains(cellReference) ){
            // this cell is outside of the formula cell area. so we just return all the cell reference targets `as is`
            return cellReferenceTargets;
        }
        // the source cell reference is inside parent formula area. so let's find target cells related to particular transformation
        // we'll iterate through all target cell references and find all the ones which belong to the target formula area
        List<CellRef> relevantCellRefs = findRelevantCellReferences(cellReferenceTargets, formulaTargetAreaRef);
        return relevantCellRefs;
    }

    private List<CellRef> findRelevantCellReferences(List<CellRef> cellReferenceTargets, AreaRef targetFormulaArea) {
        List<CellRef> relevantCellRefs = new ArrayList<>(cellReferenceTargets.size());
        for(CellRef targetCellRef: cellReferenceTargets){
            if( targetFormulaArea.contains(targetCellRef)){
                relevantCellRefs.add(targetCellRef);
            }
        }
        return relevantCellRefs;
    }

}
