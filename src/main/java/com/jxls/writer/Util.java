package com.jxls.writer;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author Leonid Vysochyn
 *         Date: 2/3/12 4:26 PM
 */
public class Util {
    private static final String regexCellRef = "([a-zA-Z]+[a-zA-Z0-9]*![a-zA-Z]+[0-9]+|[a-zA-Z]+[0-9]+|'[^?\\\\/:'*]+'![a-zA-Z]+[0-9]+)";
    private static final Pattern regexCellRefPattern = Pattern.compile(regexCellRef);

    public static List<String> getFormulaCellRefs(String formula){
        List<String> cellRefs = new ArrayList<String>();
        if( formula != null ){
            Matcher cellRefMatcher = regexCellRefPattern.matcher(formula);
            while(cellRefMatcher.find()){
                cellRefs.add(cellRefMatcher.group());
            }
        }
        return cellRefs;
    }

    public static String createTargetCellRef(List<Pos> targetCellDataList) {
        String resultRef = "";
        if( targetCellDataList != null ){
            for (Iterator<Pos> iterator = targetCellDataList.iterator(); iterator.hasNext(); ) {
                Pos pos = iterator.next();
                resultRef += pos.getCellName();
                if(iterator.hasNext()){
                    resultRef += ",";
                }
            }
        }
        return resultRef;
    }

}
