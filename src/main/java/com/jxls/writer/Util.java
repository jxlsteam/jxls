package com.jxls.writer;

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

    public static String createTargetCellRef(List<Pos> targetCellDataList) {
        String resultRef = "";
        if( targetCellDataList != null && !targetCellDataList.isEmpty()){
            List<String> cellRefs = new ArrayList<String>();
            boolean rowRange = true;
            boolean colRange = true;
            Iterator<Pos> iterator = targetCellDataList.iterator();
            Pos firstPos = iterator.next();
            cellRefs.add(firstPos.getCellName());
            String sheetName = firstPos.getSheetName();
            int row = firstPos.getRow();
            int col = firstPos.getCol();
            for (; iterator.hasNext(); ) {
                Pos pos = iterator.next();
                if( (rowRange || colRange) && !pos.getSheetName().equals(sheetName) ){
                    rowRange = false;
                    colRange = false;
                }
                if( rowRange && !(pos.getRow() - row == 1 && pos.getCol() == col )){
                    rowRange = false;
                }
                if( colRange && !(pos.getCol() - col == 1 && pos.getRow() == row)){
                    colRange = false;
                }
                sheetName = pos.getSheetName();
                row = pos.getRow();
                col = pos.getCol();
                cellRefs.add(pos.getCellName());
//                resultRef += pos.getCellName();
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

    public static List<List<Pos>> groupByRanges(List<Pos> posList, int targetRangeCount) {
        List<List<Pos>> colRanges = groupByColRange(posList);
        if( targetRangeCount == 0 || colRanges.size() == targetRangeCount) return colRanges;
        List<List<Pos>> rowRanges = groupByRowRange(posList);
        if( rowRanges.size() == targetRangeCount ) return rowRanges;
        else return colRanges;
    }

    public static List<List<Pos>> groupByColRange(List<Pos> posList) {
        List<List<Pos>> rangeList = new ArrayList<List<Pos>>();
        if(posList == null || posList.size() == 0){
            return rangeList;
        }
        List<Pos> posListCopy = new ArrayList<Pos>(posList);
        Collections.sort(posListCopy, new PosColPrecedenceComparator());

        String sheetName = posListCopy.get(0).getSheetName();
        int row = posListCopy.get(0).getRow();
        int col = posListCopy.get(0).getCol();
        List<Pos> currentRange = new ArrayList<Pos>();
        currentRange.add(posListCopy.get(0));
        boolean rangeComplete = false;
        for (int i = 1; i < posListCopy.size(); i++) {
            Pos pos = posListCopy.get(i);
            if(!pos.getSheetName().equals( sheetName ) ){
                rangeComplete = true;
            }else{
                int rowDelta = pos.getRow() - row;
                int colDelta = pos.getCol() - col;
                if(rowDelta == 1 && colDelta == 0){
                        currentRange.add( pos );
                }else {
                    rangeComplete = true;
                }
            }
            sheetName = pos.getSheetName();
            row = pos.getRow();
            col = pos.getCol();
            if( rangeComplete ){
                rangeList.add(currentRange);
                currentRange = new ArrayList<Pos>();
                currentRange.add(pos);
                rangeComplete = false;
            }
        }
        rangeList.add(currentRange);
        return rangeList;
    }

    public static List<List<Pos>> groupByRowRange(List<Pos> posList) {
        List<List<Pos>> rangeList = new ArrayList<List<Pos>>();
        if(posList == null || posList.size() == 0){
            return rangeList;
        }
        List<Pos> posListCopy = new ArrayList<Pos>(posList);
        Collections.sort(posListCopy, new PosRowPrecedenceComparator());

        String sheetName = posListCopy.get(0).getSheetName();
        int row = posListCopy.get(0).getRow();
        int col = posListCopy.get(0).getCol();
        List<Pos> currentRange = new ArrayList<Pos>();
        currentRange.add(posListCopy.get(0));
        boolean rangeComplete = false;
        for (int i = 1; i < posListCopy.size(); i++) {
            Pos pos = posListCopy.get(i);
            if(!pos.getSheetName().equals(sheetName) ){
                rangeComplete = true;
            }else{
                int rowDelta = pos.getRow() - row;
                int colDelta = pos.getCol() - col;
                if(colDelta == 1 && rowDelta == 0){
                        currentRange.add( pos );
                }else {
                    rangeComplete = true;
                }
            }
            sheetName = pos.getSheetName();
            row = pos.getRow();
            col = pos.getCol();
            if( rangeComplete ){
                rangeList.add(currentRange);
                currentRange = new ArrayList<Pos>();
                currentRange.add(pos);
                rangeComplete = false;
            }
        }
        rangeList.add(currentRange);
        return rangeList;
    }

}
