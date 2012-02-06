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
    /** The character ($) that signifies a row or column value is absolute instead of relative */
    static final char ABSOLUTE_REFERENCE_MARKER = '$';
    /** The character (!) that separates sheet names from cell references */
    private static final char SHEET_NAME_DELIMITER = '!';
    /** The character (') used to quote sheet names when they contain special characters */
    private static final char SPECIAL_NAME_DELIMITER = '\'';

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

    /**
     * takes in a column reference portion of a CellRef and converts it from
     * ALPHA-26 number format to 0-based base 10.
     * 'A' -> 0
     * 'Z' -> 25
     * 'AA' -> 26
     * 'IV' -> 255
     * @return zero based column index
     */
    public static int convertColStringToIndex(String ref) {

        int pos = 0;
        int retval=0;
        for (int k = ref.length()-1; k >= 0; k--) {
            char thechar = ref.charAt(k);
            if (thechar == ABSOLUTE_REFERENCE_MARKER) {
                if (k != 0) {
                    throw new IllegalArgumentException("Bad col ref format '" + ref + "'");
                }
                break;
            }
            // Character.getNumericValue() returns the values
            //  10-35 for the letter A-Z
            int shift = (int)Math.pow(26, pos);
            retval += (Character.getNumericValue(thechar)-9) * shift;
            pos++;
        }
        return retval-1;
    }

    public static String[] separateRefParts(String reference) {
        int plingPos = reference.lastIndexOf(SHEET_NAME_DELIMITER);
        String sheetName = parseSheetName(reference, plingPos);
        int start = plingPos+1;

        int length = reference.length();


        int loc = start;
        // skip initial dollars
        if (reference.charAt(loc)== ABSOLUTE_REFERENCE_MARKER) {
            loc++;
        }
        // step over column name chars until first digit (or dollars) for row number.
        for (; loc < length; loc++) {
            char ch = reference.charAt(loc);
            if (Character.isDigit(ch) || ch == ABSOLUTE_REFERENCE_MARKER) {
                break;
            }
        }
        return new String[] {
                sheetName,
                reference.substring(start,loc),
                reference.substring(loc),
        };
    }

    public static String parseSheetName(String reference, int indexOfSheetNameDelimiter) {
        if(indexOfSheetNameDelimiter < 0) {
            return null;
        }

        boolean isQuoted = reference.charAt(0) == SPECIAL_NAME_DELIMITER;
        if(!isQuoted) {
            return reference.substring(0, indexOfSheetNameDelimiter);
        }
        int lastQuotePos = indexOfSheetNameDelimiter-1;
        if(reference.charAt(lastQuotePos) != SPECIAL_NAME_DELIMITER) {
            throw new RuntimeException("Mismatched quotes: (" + reference + ")");
        }

        // TODO - refactor cell reference parsing logic to one place.
        // Current known incarnations:
        //   FormulaParser.GetName()
        //   CellReference.parseSheetName() (here)
        //   AreaReference.separateAreaRefs()
        //   SheetNameFormatter.format() (inverse)

        StringBuffer sb = new StringBuffer(indexOfSheetNameDelimiter);

        for(int i=1; i<lastQuotePos; i++) { // Note boundaries - skip outer quotes
            char ch = reference.charAt(i);
            if(ch != SPECIAL_NAME_DELIMITER) {
                sb.append(ch);
                continue;
            }
            if(i < lastQuotePos) {
                if(reference.charAt(i+1) == SPECIAL_NAME_DELIMITER) {
                    // two consecutive quotes is the escape sequence for a single one
                    i++; // skip this and keep parsing the special name
                    sb.append(ch);
                    continue;
                }
            }
            throw new RuntimeException("Bad sheet name quote escaping: (" + reference + ")");
        }
        return sb.toString();
    }
}
