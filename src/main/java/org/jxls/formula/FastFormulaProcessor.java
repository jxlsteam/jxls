package org.jxls.formula;

import org.jxls.common.CellData;
import org.jxls.common.CellRef;
import org.jxls.transform.Transformer;
import org.jxls.util.CellRefUtil;
import org.jxls.util.Util;

import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Fast formula processor implementation.
 * It works correctly in 90% of cases and is much more faster than {@link StandardFormulaProcessor}.
 */
public class FastFormulaProcessor implements FormulaProcessor {
    @Override
    public void processAreaFormulas(Transformer transformer) {
        Set<CellData> formulaCells = transformer.getFormulaCells();
        for (CellData formulaCellData : formulaCells) {
            List<String> formulaCellRefs = Util.getFormulaCellRefs(formulaCellData.getFormula());
            List<String> jointedCellRefs = Util.getJointedCellRefs(formulaCellData.getFormula());
            List<CellRef> targetFormulaCells = formulaCellData.getTargetPos();
            Map<CellRef, List<CellRef>> targetCellRefMap = new HashMap<CellRef, List<CellRef>>();
            Map<String, List<CellRef>> jointedCellRefMap = new HashMap<String, List<CellRef>>();
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
                boolean isFormulaCellRefsEmpty = true;
                for (Map.Entry<CellRef, List<CellRef>> cellRefEntry : targetCellRefMap.entrySet()) {
                    List<CellRef> targetCells = cellRefEntry.getValue();
                    if( targetCells.isEmpty() ) {
                        continue;
                    }
                    isFormulaCellRefsEmpty = false;
                    String replacementString;
                    if( formulaCellData.getFormulaStrategy() == CellData.FormulaStrategy.BY_COLUMN ){
                        List<CellRef> targetCellRefs = Util.createTargetCellRefListByColumn(targetFormulaCellRef, targetCells, usedCellRefs);
                        usedCellRefs.addAll(targetCellRefs);
                        replacementString = Util.createTargetCellRef(targetCellRefs);
                    }else if( targetCells.size() == targetFormulaCells.size() ){
                            CellRef targetCellRefCellRef = targetCells.get(i);
                            replacementString = targetCellRefCellRef.getCellName();
                    }else{
                            List< List<CellRef> > rangeList = Util.groupByRanges(targetCells, targetFormulaCells.size());
                            if( rangeList.size() == targetFormulaCells.size() ){
                                List<CellRef> range = rangeList.get(i);
                                replacementString = Util.createTargetCellRef(range);
                            }else{
                                replacementString = Util.createTargetCellRef(targetCells);
                            }
                    }
                    targetFormulaString = targetFormulaString.replaceAll(Util.regexJointedLookBehind + Util.sheetNameRegex(cellRefEntry) + Pattern.quote(cellRefEntry.getKey().getCellName()), Matcher.quoteReplacement(replacementString));
                }
                for (Map.Entry<String, List<CellRef>> jointedCellRefEntry : jointedCellRefMap.entrySet()) {
                    List<CellRef> targetCellRefList = jointedCellRefEntry.getValue();
                    if( targetCellRefList.isEmpty() ) {
                        continue;
                    }
                    isFormulaCellRefsEmpty = false;
                    List< List<CellRef> > rangeList = Util.groupByRanges(targetCellRefList, targetFormulaCells.size());
                    String replacementString;
                    if( rangeList.size() == targetFormulaCells.size() ){
                        List<CellRef> range = rangeList.get(i);
                        replacementString = Util.createTargetCellRef(range);
                    }else{
                        replacementString = Util.createTargetCellRef(targetCellRefList);
                    }
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
}
