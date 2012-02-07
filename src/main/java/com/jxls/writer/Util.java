package com.jxls.writer;

import java.util.*;
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

    public static List<List<Pos>> groupByRanges(List<Pos> posList) {
        List<List<Pos>> rangeList = new ArrayList<List<Pos>>();
        if(posList == null || posList.size() == 0){
            return rangeList;
        }
        List<Pos> posListCopy = new ArrayList<Pos>(posList);
        Collections.sort(posListCopy);
        int sheet = posListCopy.get(0).getSheet();
        int row = posListCopy.get(0).getRow();
        int col = posListCopy.get(0).getCol();
        List<Pos> currentRange = new ArrayList<Pos>();
        currentRange.add(posListCopy.get(0));
        boolean rangeComplete = false;
        boolean rowRange = false;
        boolean colRange = false;
        for (int i = 1; i < posListCopy.size(); i++) {
            Pos pos = posListCopy.get(i);
            if(pos.getSheet() != sheet ){
                rangeComplete = true;
            }else{
                int rowDelta = pos.getRow() - row;
                int colDelta = pos.getCol() - col;
                if(rowDelta == 1 && colDelta == 0){
                    if( !colRange ){
                        rowRange = true;
                        currentRange.add( pos );
                    }else{
                        rangeComplete = true;
                    }
                }else if( colDelta == 1 && rowDelta == 0){
                    if( !rowRange ){
                        colRange = true;
                        currentRange.add( pos );
                    }else{
                        rangeComplete = true;
                    }
                }else{
                    rangeComplete = true;
                }
            }
            sheet = pos.getSheet();
            row = pos.getRow();
            col = pos.getCol();
            if( rangeComplete ){
                rangeList.add(currentRange);
                currentRange = new ArrayList<Pos>();
                currentRange.add(pos);
                rowRange = false;
                colRange = false;
                rangeComplete = false;
            }
        }
        rangeList.add(currentRange);
        return rangeList;
    }
}
