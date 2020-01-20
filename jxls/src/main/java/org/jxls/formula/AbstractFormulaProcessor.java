package org.jxls.formula;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.jxls.area.Area;
import org.jxls.common.CellData;
import org.jxls.common.CellRef;
import org.jxls.transform.Transformer;
import org.jxls.util.Util;

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
        List<String> formulaCellRefs = Util.getFormulaCellRefs(formulaCellData.getFormula());
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
        List<String> jointedCellRefs = Util.getJointedCellRefs(formulaCellData.getFormula());
        // for each jointed cell ref build a list of cells into which individual cell names were transformed to
        for (String jointedCellRef : jointedCellRefs) {
            List<String> nestedCellRefs = Util.getCellRefsFromJointedCellRef(jointedCellRef);
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

    @Deprecated // TODO MW to Leonid: What's our plan when to kick this method?
    @Override
    public void processAreaFormulas(Transformer transformer) {
        processAreaFormulas(transformer, null);
    }
}
