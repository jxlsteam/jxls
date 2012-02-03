package com.jxls.writer;

import java.util.ArrayList;
import java.util.List;

/**
 * @author Leonid Vysochyn
 *         Date: 2/3/12 4:26 PM
 */
public class Util {
    public static List<String> getFormulaCellRefs(String formula){
        List<String> cellRefs = new ArrayList<String>();
        return cellRefs;
    }

    public static String createTargetCellRef(List<Pos> targetCellDataList) {
        String resultRef = "";
        for (int i = 0; i < targetCellDataList.size(); i++) {
            Pos pos = targetCellDataList.get(i);
            resultRef += pos.getCellName() + ",";
        }
        return resultRef;
    }
}
