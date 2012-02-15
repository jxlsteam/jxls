package com.jxls.writer.util;

import com.jxls.writer.common.CellRef;
import com.jxls.writer.common.CellRefColPrecedenceComparator;
import com.jxls.writer.common.CellRefRowPrecedenceComparator;

import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author Leonid Vysochyn
 *         Date: 2/3/12 4:26 PM
 */
public class Util {
    public static final String regexJointedLookBehind = "(?<!U_\\([^)]{0,100})";
    private static final String regexCellRef = "([a-zA-Z]+[a-zA-Z0-9]*![a-zA-Z]+[0-9]+|[a-zA-Z]+[0-9]+|'[^?\\\\/:'*]+'![a-zA-Z]+[0-9]+)";
    private static final String regexCellRefExcludingJointed = regexJointedLookBehind + regexCellRef;
    private static final Pattern regexCellRefExcludingJointedPattern = Pattern.compile(regexCellRefExcludingJointed);
    private static final Pattern regexCellRefPattern = Pattern.compile(regexCellRef);
    private static final String regexJointedCellRef = "U_\\([^\\)]+\\)";
    private static final Pattern regexJointedCellRefPattern = Pattern.compile(regexJointedCellRef);

    public static List<String> getFormulaCellRefs(String formula){
        return getStringPartsByPattern(formula, regexCellRefExcludingJointedPattern);
    }

    private static List<String> getStringPartsByPattern(String str, Pattern pattern) {
        List<String> cellRefs = new ArrayList<String>();
        if( str != null ){
            Matcher cellRefMatcher = pattern.matcher(str);
            while(cellRefMatcher.find()){
                cellRefs.add(cellRefMatcher.group());
            }
        }
        return cellRefs;
    }

    public static List<String> getJointedCellRefs(String formula){
        return getStringPartsByPattern(formula, regexJointedCellRefPattern);
    }
    
    public static List<String> getCellRefsFromJointedCellRef(String jointedCellRef){
        return getStringPartsByPattern(jointedCellRef, regexCellRefPattern);
    }
    
    public static boolean formulaContainsJointedCellRef(String formula){
        return regexJointedCellRefPattern.matcher(formula).find();
    }

    public static String createTargetCellRef(List<CellRef> targetCellDataList) {
        String resultRef = "";
        if( targetCellDataList != null && !targetCellDataList.isEmpty()){
            List<String> cellRefs = new ArrayList<String>();
            boolean rowRange = true;
            boolean colRange = true;
            Iterator<CellRef> iterator = targetCellDataList.iterator();
            CellRef firstCellRef = iterator.next();
            cellRefs.add(firstCellRef.getCellName());
            String sheetName = firstCellRef.getSheetName();
            int row = firstCellRef.getRow();
            int col = firstCellRef.getCol();
            for (; iterator.hasNext(); ) {
                CellRef cellRef = iterator.next();
                if( (rowRange || colRange) && !cellRef.getSheetName().equals(sheetName) ){
                    rowRange = false;
                    colRange = false;
                }
                if( rowRange && !(cellRef.getRow() - row == 1 && cellRef.getCol() == col )){
                    rowRange = false;
                }
                if( colRange && !(cellRef.getCol() - col == 1 && cellRef.getRow() == row)){
                    colRange = false;
                }
                sheetName = cellRef.getSheetName();
                row = cellRef.getRow();
                col = cellRef.getCol();
                cellRefs.add(cellRef.getCellName());
//                resultRef += cellRef.getCellName();
//                if(iterator.hasNext()){
//                    resultRef += ",";
//                }
            }
            if( (rowRange || colRange) && cellRefs.size() > 1 ){
                resultRef = cellRefs.get(0) + ":" + cellRefs.get( cellRefs.size() -1 );
            }else{
                resultRef = joinStrings(cellRefs, ",");
            }
        }
        return resultRef;
    }
    
    public static String joinStrings(List<String> strings, String separator){
        StringBuilder sb = new StringBuilder();
        String sep = "";
        for(String s: strings) {
            sb.append(sep).append(s);
            sep = separator;
        }
        return sb.toString();
    }

    public static List<List<CellRef>> groupByRanges(List<CellRef> cellRefList, int targetRangeCount) {
        List<List<CellRef>> colRanges = groupByColRange(cellRefList);
        if( targetRangeCount == 0 || colRanges.size() == targetRangeCount) return colRanges;
        List<List<CellRef>> rowRanges = groupByRowRange(cellRefList);
        if( rowRanges.size() == targetRangeCount ) return rowRanges;
        else return colRanges;
    }

    public static List<List<CellRef>> groupByColRange(List<CellRef> cellRefList) {
        List<List<CellRef>> rangeList = new ArrayList<List<CellRef>>();
        if(cellRefList == null || cellRefList.size() == 0){
            return rangeList;
        }
        List<CellRef> cellRefListCopy = new ArrayList<CellRef>(cellRefList);
        Collections.sort(cellRefListCopy, new CellRefColPrecedenceComparator());

        String sheetName = cellRefListCopy.get(0).getSheetName();
        int row = cellRefListCopy.get(0).getRow();
        int col = cellRefListCopy.get(0).getCol();
        List<CellRef> currentRange = new ArrayList<CellRef>();
        currentRange.add(cellRefListCopy.get(0));
        boolean rangeComplete = false;
        for (int i = 1; i < cellRefListCopy.size(); i++) {
            CellRef cellRef = cellRefListCopy.get(i);
            if(!cellRef.getSheetName().equals( sheetName ) ){
                rangeComplete = true;
            }else{
                int rowDelta = cellRef.getRow() - row;
                int colDelta = cellRef.getCol() - col;
                if(rowDelta == 1 && colDelta == 0){
                        currentRange.add(cellRef);
                }else {
                    rangeComplete = true;
                }
            }
            sheetName = cellRef.getSheetName();
            row = cellRef.getRow();
            col = cellRef.getCol();
            if( rangeComplete ){
                rangeList.add(currentRange);
                currentRange = new ArrayList<CellRef>();
                currentRange.add(cellRef);
                rangeComplete = false;
            }
        }
        rangeList.add(currentRange);
        return rangeList;
    }

    public static List<List<CellRef>> groupByRowRange(List<CellRef> cellRefList) {
        List<List<CellRef>> rangeList = new ArrayList<List<CellRef>>();
        if(cellRefList == null || cellRefList.size() == 0){
            return rangeList;
        }
        List<CellRef> cellRefListCopy = new ArrayList<CellRef>(cellRefList);
        Collections.sort(cellRefListCopy, new CellRefRowPrecedenceComparator());

        String sheetName = cellRefListCopy.get(0).getSheetName();
        int row = cellRefListCopy.get(0).getRow();
        int col = cellRefListCopy.get(0).getCol();
        List<CellRef> currentRange = new ArrayList<CellRef>();
        currentRange.add(cellRefListCopy.get(0));
        boolean rangeComplete = false;
        for (int i = 1; i < cellRefListCopy.size(); i++) {
            CellRef cellRef = cellRefListCopy.get(i);
            if(!cellRef.getSheetName().equals(sheetName) ){
                rangeComplete = true;
            }else{
                int rowDelta = cellRef.getRow() - row;
                int colDelta = cellRef.getCol() - col;
                if(colDelta == 1 && rowDelta == 0){
                        currentRange.add(cellRef);
                }else {
                    rangeComplete = true;
                }
            }
            sheetName = cellRef.getSheetName();
            row = cellRef.getRow();
            col = cellRef.getCol();
            if( rangeComplete ){
                rangeList.add(currentRange);
                currentRange = new ArrayList<CellRef>();
                currentRange.add(cellRef);
                rangeComplete = false;
            }
        }
        rangeList.add(currentRange);
        return rangeList;
    }

}
