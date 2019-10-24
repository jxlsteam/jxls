package org.jxls.util;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.jxls.common.JxlsException;

/**
 * This is a class to convert Excel cell names to (sheet, row, col) representations and vice versa.
 * The current code is taken from Apache POI CellReference class ( http://poi.apache.org/apidocs/org/apache/poi/ss/util/CellReference.html ).
 * 
 * @author Leonid Vysochyn
 */
public class CellRefUtil {
    private static final char DELIMITER = '\'';
    /** Matches a single cell ref with no absolute ('$') markers */
    private static final Pattern CELL_REF_PATTERN = Pattern.compile("([A-Za-z]+)([0-9]+)");
    /** The character (!) that separates sheet names from cell references */
    public static final char SHEET_NAME_DELIMITER = '!';
    /** The character (:) that separates the two cell references in a multi-cell area reference */
    private static final char CELL_DELIMITER = ':';
    /** The character ($) that signifies a row or column value is absolute instead of relative */
    public static final char ABSOLUTE_REFERENCE_MARKER = '$';
    /** The character (') used to quote sheet names when they contain special characters */
    private static final char SPECIAL_NAME_DELIMITER = '\'';

    /**
     * Takes in a 0-based base-10 column and returns a ALPHA-26 representation. e.g. column #3 -&gt; D
     * 
     * @param col -
     * @return -
     */
    public static String convertNumToColString(int col) {
        String colRef = "";
        int excelColNum = col + 1; // Excel counts column A as the 1st column, we treat it as the 0th one.
        int colRemain = excelColNum;
        while (colRemain > 0) {
            int thisPart = colRemain % 26;
            if (thisPart == 0) {
                thisPart = 26;
            }
            colRemain = (colRemain - thisPart) / 26;
            char colChar = (char) (thisPart + 64); // The letter A is at 65
            colRef = colChar + colRef;
        }
        return colRef;
    }

    public static void appendFormat(StringBuilder out, String rawSheetName) {
        boolean needsQuotes = needsDelimiting(rawSheetName);
        if (needsQuotes) {
            out.append(DELIMITER);
            appendAndEscape(out, rawSheetName);
            out.append(DELIMITER);
        } else {
            out.append(rawSheetName);
        }
    }

    static void appendAndEscape(StringBuilder sb, String rawSheetName) {
        int len = rawSheetName.length();
        for (int i = 0; i < len; i++) {
            char ch = rawSheetName.charAt(i);
            if (ch == DELIMITER) {
                // single quotes (') are encoded as ('')
                sb.append(DELIMITER);
            }
            sb.append(ch);
        }
    }

    static boolean needsDelimiting(String rawSheetName) {
        int len = rawSheetName.length();
        if (len < 1) {
            throw new JxlsException("Zero length string is an invalid sheet name");
        }
        if (Character.isDigit(rawSheetName.charAt(0))) {
            // sheet name with digit in the first position always requires delimiting
            return true;
        }
        for (int i = 0; i < len; i++) {
            char ch = rawSheetName.charAt(i);
            if (isSpecialChar(ch)) {
                return true;
            }
        }
        if (Character.isLetter(rawSheetName.charAt(0)) && Character.isDigit(rawSheetName.charAt(len - 1)) // note - values like "A$1:$C$20" don't get this far
                && nameLooksLikePlainCellReference(rawSheetName)) {
            return true;
        }
        return nameLooksLikeBooleanLiteral(rawSheetName);
    }

    /**
     * Note - this method assumes the specified rawSheetName has only letters and digits.  It
     * cannot be used to match absolute or range references (using the dollar or colon char).
     * 
     * <p>Some notable cases:
     *    <blockquote><table border="0" cellpadding="1" cellspacing="0"
     *                 summary="Notable cases.">
     *      <tr><th>Input&nbsp;</th><th>Result&nbsp;</th><th>Comments</th></tr>
     *      <tr><td>"A1"&nbsp;&nbsp;</td><td>true</td><td>&nbsp;</td></tr>
     *      <tr><td>"a111"&nbsp;&nbsp;</td><td>true</td><td>&nbsp;</td></tr>
     *      <tr><td>"AA"&nbsp;&nbsp;</td><td>false</td><td>&nbsp;</td></tr>
     *      <tr><td>"aa1"&nbsp;&nbsp;</td><td>true</td><td>&nbsp;</td></tr>
     *      <tr><td>"A1A"&nbsp;&nbsp;</td><td>false</td><td>&nbsp;</td></tr>
     *      <tr><td>"A1A1"&nbsp;&nbsp;</td><td>false</td><td>&nbsp;</td></tr>
     *      <tr><td>"A$1:$C$20"&nbsp;&nbsp;</td><td>false</td><td>Not a plain cell reference</td></tr>
     *      <tr><td>"SALES20080101"&nbsp;&nbsp;</td><td>true</td>
     *      		<td>Still needs delimiting even though well out of range</td></tr>
     *    </table></blockquote></p>
     *
     * @return <code>true</code> if there is any possible ambiguity that the specified rawSheetName
     * could be interpreted as a valid cell name.
     */
    static boolean nameLooksLikePlainCellReference(String rawSheetName) {
        Matcher matcher = CELL_REF_PATTERN.matcher(rawSheetName);
        if (!matcher.matches()) {
            return false;
        }

        // rawSheetName == "Sheet1" gets this far.
        String lettersPrefix = matcher.group(1);
        String numbersSuffix = matcher.group(2);
        return cellReferenceIsWithinRange(lettersPrefix, numbersSuffix);
    }

    /**
     * Used to decide whether sheet names like 'AB123' need delimiting due to the fact that they
     * look like cell references.
     * 
     * <p>This code is currently being used for translating formulas represented with <code>Ptg</code>
     * tokens into human readable text form.  In formula expressions, a sheet name always has a
     * trailing '!' so there is little chance for ambiguity.  It doesn't matter too much what this
     * method returns but it is worth noting the likely consumers of these formula text strings:
     * <ol>
     * <li>POI's own formula parser</li>
     * <li>Visual reading by human</li>
     * <li>VBA automation entry into Excel cell contents e.g.  ActiveCell.Formula = "=c64!A1"</li>
     * <li>Manual entry into Excel cell contents</li>
     * <li>Some third party formula parser</li>
     * </ol></p>
     *
     * <p>At the time of writing, POI's formula parser tolerates cell-like sheet names in formulas
     * with or without delimiters.  The same goes for Excel(2007), both manual and automated entry.</p>
     * 
     * <p>For better or worse this implementation attempts to replicate Excel's formula renderer.
     * Excel uses range checking on the apparent 'row' and 'column' components.  Note however that
     * the maximum sheet size varies across versions.</p>
     */
    private static boolean cellReferenceIsWithinRange(String lettersPrefix, String numbersSuffix) {
        return cellReferenceIsWithinRange(lettersPrefix, numbersSuffix, 0x0100, 0x10000);
    }

    /**
     * Used to decide whether a name of the form "[A-Z]*[0-9]*" that appears in a formula can be
     * interpreted as a cell reference.  Names of that form can be also used for sheets and/or
     * named ranges, and in those circumstances, the question of whether the potential cell
     * reference is valid (in range) becomes important.
     * 
     * <p>Note - that the maximum sheet size varies across Excel versions:</p>
     * 
     * <blockquote><table border="0" cellpadding="1" cellspacing="0"
     *                 summary="Notable cases.">
     *   <tr><th>Version&nbsp;&nbsp;</th><th>File Format&nbsp;&nbsp;</th>
     *   	<th>Last Column&nbsp;&nbsp;</th><th>Last Row</th></tr>
     *   <tr><td>97-2003</td><td>BIFF8</td><td>"IV" (2^8)</td><td>65536 (2^14)</td></tr>
     *   <tr><td>2007</td><td>BIFF12</td><td>"XFD" (2^14)</td><td>1048576 (2^20)</td></tr>
     * </table></blockquote>
     * POI currently targets BIFF8 (Excel 97-2003), so the following behaviour can be observed for
     * this method:
     * <blockquote><table border="0" cellpadding="1" cellspacing="0"
     *                 summary="Notable cases.">
     *   <tr><th>Input&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</th>
     *       <th>Result&nbsp;</th></tr>
     *   <tr><td>"A", "1"</td><td>true</td></tr>
     *   <tr><td>"a", "111"</td><td>true</td></tr>
     *   <tr><td>"A", "65536"</td><td>true</td></tr>
     *   <tr><td>"A", "65537"</td><td>false</td></tr>
     *   <tr><td>"iv", "1"</td><td>true</td></tr>
     *   <tr><td>"IW", "1"</td><td>false</td></tr>
     *   <tr><td>"AAA", "1"</td><td>false</td></tr>
     *   <tr><td>"a", "111"</td><td>true</td></tr>
     *   <tr><td>"Sheet", "1"</td><td>false</td></tr>
     * </table></blockquote>
     *
     * @param colStr a string of only letter characters
     * @param rowStr a string of only digit characters
     * @param lastColumnIndex -
     * @param lastRowIndex -
     * @return <code>true</code> if the row and col parameters are within range of a BIFF8 spreadsheet.
     */
    public static boolean cellReferenceIsWithinRange(String colStr, String rowStr, int lastColumnIndex, int lastRowIndex) {
        if (!isColumnWithnRange(colStr, lastColumnIndex)) {
            return false;
        }
        return isRowWithnRange(rowStr, lastRowIndex);
    }

    public static boolean isColumnWithnRange(String colStr, int lastColumnIndex) {
        String lastCol = convertNumToColString(lastColumnIndex);
        int lastColLength = lastCol.length();
        int numberOfLetters = colStr.length();
        if (numberOfLetters > lastColLength) {
            // "Sheet1" case etc
            return false; // that was easy
        }
        return !(numberOfLetters == lastColLength && colStr.toUpperCase().compareTo(lastCol) > 0);
    }

    public static boolean isRowWithnRange(String rowStr, int lastRowIndex) {
        int rowNum = Integer.parseInt(rowStr);
        if (rowNum < 0) {
            throw new IllegalStateException("Invalid rowStr '" + rowStr + "'.");
        }
        if (rowNum == 0) {
            // execution gets here because caller does first pass of discriminating
            // potential cell references using a simplistic regex pattern.
            return false;
        }
        return rowNum <= lastRowIndex;
    }


    static boolean nameLooksLikeBooleanLiteral(String rawSheetName) {
        switch(rawSheetName.charAt(0)) {
            case 'T':
            case 't':
                return "TRUE".equalsIgnoreCase(rawSheetName);
            case 'F':
            case 'f':
                return "FALSE".equalsIgnoreCase(rawSheetName);
        }
        return false;
    }

    /**
     * @return <code>true</code> if the presence of the specified character in a sheet name would
     * require the sheet name to be delimited in formulas.  This includes every non-alphanumeric
     * character besides underscore '_' and dot '.'.
     */
    static boolean isSpecialChar(char ch) {
        // note - Character.isJavaIdentifierPart() would allow dollars '$'
        if (Character.isLetterOrDigit(ch)) {
            return false;
        }
        switch(ch) {
            case '.': // dot is OK
            case '_': // underscore is OK
                return false;
            case '\n':
            case '\r':
            case '\t':
                throw new JxlsException("Illegal character (0x" + Integer.toHexString(ch) + ") found in sheet name");
            default:
        }
        return true;
    }

    /**
     * takes in a column reference portion of a CellRef and converts it from
     * ALPHA-26 number format to 0-based base 10.
     * <pre> 'A' -&gt; 0
     * 'Z' -&gt; 25
     * 'AA' -&gt; 26
     * 'IV' -&gt; 255</pre>
     * 
     * @param ref -
     * @return zero based column index
     */
    public static int convertColStringToIndex(String ref) {
        int pos = 0;
        int retval = 0;
        for (int k = ref.length() - 1; k >= 0; k--) {
            char thechar = ref.charAt(k);
            if (thechar == ABSOLUTE_REFERENCE_MARKER) {
                if (k != 0) {
                    throw new IllegalArgumentException("Bad col ref format '" + ref + "'");
                }
                break;
            }
            // Character.getNumericValue() returns the values
            //  10-35 for the letter A-Z
            int shift = (int) Math.pow(26, pos);
            retval += (Character.getNumericValue(thechar) - 9) * shift;
            pos++;
        }
        return retval - 1;
    }

    public static String[] separateRefParts(String reference) {
        int plingPos = reference.lastIndexOf(SHEET_NAME_DELIMITER);
        String sheetName = parseSheetName(reference, plingPos);
        int start = plingPos + 1;

        int length = reference.length();

        int loc = start;
        // skip initial dollars
        if (reference.charAt(loc) == ABSOLUTE_REFERENCE_MARKER) {
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
        if (indexOfSheetNameDelimiter < 0) {
            return null;
        }

        boolean isQuoted = reference.charAt(0) == SPECIAL_NAME_DELIMITER;
        if (!isQuoted) {
            return reference.substring(0, indexOfSheetNameDelimiter);
        }
        int lastQuotePos = indexOfSheetNameDelimiter-1;
        if (reference.charAt(lastQuotePos) != SPECIAL_NAME_DELIMITER) {
            throw new JxlsException("Mismatched quotes: (" + reference + ")");
        }

        // TODO - refactor cell reference parsing logic to one place.
        // Current known incarnations:
        //   FormulaParser.GetName()
        //   CellReference.parseSheetName() (here)
        //   AreaReference.separateAreaRefs()
        //   SheetNameFormatter.format() (inverse)

        StringBuilder sb = new StringBuilder(indexOfSheetNameDelimiter);

        for(int i = 1; i < lastQuotePos; i++) { // Note boundaries - skip outer quotes
            char ch = reference.charAt(i);
            if (ch != SPECIAL_NAME_DELIMITER) {
                sb.append(ch);
                continue;
            }
            if (i < lastQuotePos && reference.charAt(i+1) == SPECIAL_NAME_DELIMITER) {
                // two consecutive quotes is the escape sequence for a single one
                i++; // skip this and keep parsing the special name
                sb.append(ch);
                continue;
            }
            throw new JxlsException("Bad sheet name quote escaping: (" + reference + ")");
        }
        return sb.toString();
    }

    /**
     * Separates Area refs in two parts and returns them as separate elements in a String array,
     * each qualified with the sheet name (if present)
     *
     * @param reference -
     * @return array with one or two elements. never <code>null</code>
     */
    public static String[] separateAreaRefs(String reference) {
        int len = reference.length();
        int delimiterPos = -1;
        boolean insideDelimitedName = false;
        for (int i = 0; i < len; i++) {
            switch (reference.charAt(i)) {
                case CELL_DELIMITER:
                    if (!insideDelimitedName) {
                        if (delimiterPos >= 0) {
                            throw new IllegalArgumentException("More than one cell delimiter '"
                                    + CELL_DELIMITER + "' appears in area reference '" + reference + "'");
                        }
                        delimiterPos = i;
                    }
                default:
                    continue;
                case SPECIAL_NAME_DELIMITER:
                    // fall through
            }
            if (!insideDelimitedName) {
                insideDelimitedName = true;
                continue;
            }

            if (i >= len - 1) {
                // reference ends with the delimited name.
                // Assume names like: "Sheet1!'A1'" are never legal.
                throw new IllegalArgumentException("Area reference '" + reference
                        + "' ends with special name delimiter '"  + SPECIAL_NAME_DELIMITER + "'");
            }
            if (reference.charAt(i + 1) == SPECIAL_NAME_DELIMITER) {
                // two consecutive quotes is the escape sequence for a single one
                i++; // skip this and keep parsing the special name
            } else {
                // this is the end of the delimited name
                insideDelimitedName = false;
            }
        }
        if (delimiterPos < 0) {
            return new String[] { reference };
        }

        String partA = reference.substring(0, delimiterPos);
        String partB = reference.substring(delimiterPos+1);
        if (partB.indexOf(SHEET_NAME_DELIMITER) >= 0) {
            throw new JxlsException("Unexpected " + SHEET_NAME_DELIMITER
                    + " in second cell reference of '" + reference + "'");
        }

        int plingPos = partA.lastIndexOf(SHEET_NAME_DELIMITER);
        if (plingPos < 0) {
            return new String[] { partA, partB, };
        }

        String sheetName = partA.substring(0, plingPos + 1); // +1 to include delimiter

        return new String[] { partA, sheetName + partB, };
    }

    public static boolean isPlainColumn(String refPart) {
        for (int i = refPart.length() - 1; i >= 0; i--) {
            int ch = refPart.charAt(i);
            if (ch == '$' && i == 0) {
                continue;
            }
            if (ch < 'A' || ch > 'Z') {
                return false;
            }
        }
        return true;
    }
}
