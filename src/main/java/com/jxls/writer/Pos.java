package com.jxls.writer;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.CellReference;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author Leonid Vysochyn
 *         Date: 1/25/12 1:42 PM
 */
public class Pos {
    private static final char DELIMITER = '\'';
    /**
     * Matches a single cell ref with no absolute ('$') markers
     */
    private static final Pattern CELL_REF_PATTERN = Pattern.compile("([A-Za-z]+)([0-9]+)");

    /** The character (!) that separates sheet names from cell references */
    private static final char SHEET_NAME_DELIMITER = '!';

    /** The character ($) that signifies a row or column value is absolute instead of relative */
    private static final char ABSOLUTE_REFERENCE_MARKER = '$';

    int col;
    int row;
    int sheet;
    String sheetName;
    boolean isColAbs;
    boolean isRowAbs;

    public Pos(int sheet, int row, int col) {
        this.sheet = sheet;
        this.row = row;
        this.col = col;
    }

    public Pos(int row, int col) {
        this(0, row, col);
    }
    
    public Pos(String cellRef){
        if(cellRef.endsWith("#REF!")) {
            throw new IllegalArgumentException("Cell reference invalid: " + cellRef);
        }

        String[] parts = Util.separateRefParts(cellRef);
        sheetName = parts[0];
        String colRef = parts[1];
        if (colRef.length() < 1) {
            throw new IllegalArgumentException("Invalid Formula cell reference: '"+cellRef+"'");
        }
        isColAbs = colRef.charAt(0) == '$';
        if (isColAbs) {
            colRef=colRef.substring(1);
        }
        col = Util.convertColStringToIndex(colRef);

        String rowRef=parts[2];
        if (rowRef.length() < 1) {
            throw new IllegalArgumentException("Invalid Formula cell reference: '"+cellRef+"'");
        }
        isRowAbs = rowRef.charAt(0) == '$';
        if (isRowAbs) {
            rowRef=rowRef.substring(1);
        }
        row = Integer.parseInt(rowRef)-1; // -1 to convert 1-based to zero-based
    }

    public String getCellName(){
        StringBuffer sb = new StringBuffer(32);
        if(sheetName != null) {
            appendFormat(sb, sheetName);
            sb.append(SHEET_NAME_DELIMITER);
        }
        appendCellReference(sb);
        return sb.toString();
    }

    /**
     * Appends cell reference with '$' markers for absolute values as required.
     * Sheet name is not included.
     */
    /* package */ void appendCellReference(StringBuffer sb) {
        if(isColAbs) {
            sb.append(ABSOLUTE_REFERENCE_MARKER);
        }
        sb.append( convertNumToColString(col));
        if(isRowAbs) {
            sb.append(ABSOLUTE_REFERENCE_MARKER);
        }
        sb.append(row+1);
    }

    /**
     * Takes in a 0-based base-10 column and returns a ALPHA-26
     *  representation.
     * eg column #3 -> D
     */
    public static String convertNumToColString(int col) {
        // Excel counts column A as the 1st column, we
        //  treat it as the 0th one
        int excelColNum = col + 1;

        String colRef = "";
        int colRemain = excelColNum;

        while(colRemain > 0) {
            int thisPart = colRemain % 26;
            if(thisPart == 0) { thisPart = 26; }
            colRemain = (colRemain - thisPart) / 26;

            // The letter A is at 65
            char colChar = (char)(thisPart+64);
            colRef = colChar + colRef;
        }

        return colRef;
    }

    public static void appendFormat(StringBuffer out, String rawSheetName) {
        boolean needsQuotes = needsDelimiting(rawSheetName);
        if(needsQuotes) {
            out.append(DELIMITER);
            appendAndEscape(out, rawSheetName);
            out.append(DELIMITER);
        } else {
            out.append(rawSheetName);
        }
    }

    private static void appendAndEscape(StringBuffer sb, String rawSheetName) {
        int len = rawSheetName.length();
        for(int i=0; i<len; i++) {
            char ch = rawSheetName.charAt(i);
            if(ch == DELIMITER) {
                // single quotes (') are encoded as ('')
                sb.append(DELIMITER);
            }
            sb.append(ch);
        }
    }

    private static boolean needsDelimiting(String rawSheetName) {
        int len = rawSheetName.length();
        if(len < 1) {
            throw new RuntimeException("Zero length string is an invalid sheet name");
        }
        if(Character.isDigit(rawSheetName.charAt(0))) {
            // sheet name with digit in the first position always requires delimiting
            return true;
        }
        for(int i=0; i<len; i++) {
            char ch = rawSheetName.charAt(i);
            if(isSpecialChar(ch)) {
                return true;
            }
        }
        if(Character.isLetter(rawSheetName.charAt(0))
                && Character.isDigit(rawSheetName.charAt(len-1))) {
            // note - values like "A$1:$C$20" don't get this far
            if(nameLooksLikePlainCellReference(rawSheetName)) {
                return true;
            }
        }
        if (nameLooksLikeBooleanLiteral(rawSheetName)) {
            return true;
        }
        // Error constant literals all contain '#' and other special characters
        // so they don't get this far
        return false;
    }

    /**
     * Note - this method assumes the specified rawSheetName has only letters and digits.  It
     * cannot be used to match absolute or range references (using the dollar or colon char).
     * <p/>
     * Some notable cases:
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
     *    </table></blockquote>
     *
     * @return <code>true</code> if there is any possible ambiguity that the specified rawSheetName
     * could be interpreted as a valid cell name.
     */
    /* package */ static boolean nameLooksLikePlainCellReference(String rawSheetName) {
        Matcher matcher = CELL_REF_PATTERN.matcher(rawSheetName);
        if(!matcher.matches()) {
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
     * <p/>
     * This code is currently being used for translating formulas represented with <code>Ptg</code>
     * tokens into human readable text form.  In formula expressions, a sheet name always has a
     * trailing '!' so there is little chance for ambiguity.  It doesn't matter too much what this
     * method returns but it is worth noting the likely consumers of these formula text strings:
     * <ol>
     * <li>POI's own formula parser</li>
     * <li>Visual reading by human</li>
     * <li>VBA automation entry into Excel cell contents e.g.  ActiveCell.Formula = "=c64!A1"</li>
     * <li>Manual entry into Excel cell contents</li>
     * <li>Some third party formula parser</li>
     * </ol>
     *
     * At the time of writing, POI's formula parser tolerates cell-like sheet names in formulas
     * with or without delimiters.  The same goes for Excel(2007), both manual and automated entry.
     * <p/>
     * For better or worse this implementation attempts to replicate Excel's formula renderer.
     * Excel uses range checking on the apparent 'row' and 'column' components.  Note however that
     * the maximum sheet size varies across versions.
     * @see org.apache.poi.ss.util.CellReference
     */
    /* package */ static boolean cellReferenceIsWithinRange(String lettersPrefix, String numbersSuffix) {
        return CellReference.cellReferenceIsWithinRange(lettersPrefix, numbersSuffix, SpreadsheetVersion.EXCEL97);
    }

    private static boolean nameLooksLikeBooleanLiteral(String rawSheetName) {
        switch(rawSheetName.charAt(0)) {
            case 'T': case 't':
                return "TRUE".equalsIgnoreCase(rawSheetName);
            case 'F': case 'f':
                return "FALSE".equalsIgnoreCase(rawSheetName);
        }
        return false;
    }
    /**
     * @return <code>true</code> if the presence of the specified character in a sheet name would
     * require the sheet name to be delimited in formulas.  This includes every non-alphanumeric
     * character besides underscore '_' and dot '.'.
     */
    /* package */ static boolean isSpecialChar(char ch) {
        // note - Character.isJavaIdentifierPart() would allow dollars '$'
        if(Character.isLetterOrDigit(ch)) {
            return false;
        }
        switch(ch) {
            case '.': // dot is OK
            case '_': // underscore is OK
                return false;
            case '\n':
            case '\r':
            case '\t':
                throw new RuntimeException("Illegal character (0x"
                        + Integer.toHexString(ch) + ") found in sheet name");
        }
        return true;
    }

    public int getSheet() {
        return sheet;
    }

    public void setSheet(int sheet) {
        this.sheet = sheet;
    }

    public int getCol() {
        return col;
    }

    public void setCol(int col) {
        this.col = col;
    }

    public int getRow() {
        return row;
    }

    public void setRow(int row) {
        this.row = row;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;

        Pos pos = (Pos) o;

        if (col != pos.col) return false;
        if (row != pos.row) return false;
        if (sheet != pos.sheet) return false;

        return true;
    }

    @Override
    public int hashCode() {
        int result = col;
        result = 31 * result + row;
        result = 31 * result + sheet;
        return result;
    }

    @Override
    public String toString() {
        return "Pos{" +
                "sheet=" + sheet +
                ", row=" + row +
                ", col=" + col +
                '}';
    }
}
