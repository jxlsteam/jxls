package com.jxls.writer.transform.poi;

import com.jxls.writer.command.Context;
import com.jxls.writer.expression.ExpressionEvaluator;
import com.jxls.writer.expression.JexlExpressionEvaluator;
import org.apache.poi.ss.usermodel.*;

import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author Leonid Vysochyn
 *         Date: 1/23/12 6:39 PM
 */
public class CellData {
    protected static final String REGEX_EXPRESSION = "\\$\\{[^}]*}";
    protected static final Pattern REGEX_EXPRESSION_PATTERN = Pattern.compile(REGEX_EXPRESSION);


    RichTextString richTextString;
    boolean booleanValue;

    int cellType;
    int resultCellType;
    private Date dateValue;
    private double doubleValue;
    private String formula;
    private CellStyle style;
    private Hyperlink hyperlink;
    private byte errorValue;
    private int rowIndex;
    private int colIndex;
    private Object evaluationResult;
    private Object cellOriginalValue;

    public static CellData createCellData(Cell cell){
        CellData cellData = new CellData();
        cellData.readCell(cell);
        return cellData;
    }
    
    public Object getCellValue(){
        return cellOriginalValue;
    }
    
    public Object evaluate(Context context){
        resultCellType = cellType;
        if( richTextString != null){
            String strValue = richTextString.getString();
            StringBuffer sb = new StringBuffer();
            Matcher exprMatcher = REGEX_EXPRESSION_PATTERN.matcher(strValue);
            ExpressionEvaluator evaluator = new JexlExpressionEvaluator(context.toMap());
            String matchedString;
            String expression;
            Object lastMatchEvalResult = null;
            int matchCount = 0;
            int endOffset = 0;
            while(exprMatcher.find()){
                endOffset = exprMatcher.end();
                matchCount++;
                matchedString = exprMatcher.group();
                expression = matchedString.substring(2, matchedString.length() - 1);
                lastMatchEvalResult = evaluator.evaluate(expression);
                exprMatcher.appendReplacement(sb, Matcher.quoteReplacement( lastMatchEvalResult != null ? lastMatchEvalResult.toString() : "" ));
            }
            if( matchCount > 1 || (matchCount == 1 && endOffset < strValue.length()-1)){
                exprMatcher.appendTail(sb);
                evaluationResult = sb.toString();
            }else if(matchCount == 1){
                evaluationResult = lastMatchEvalResult;
                if(evaluationResult instanceof Number){
                    resultCellType = Cell.CELL_TYPE_NUMERIC;
                }else if(evaluationResult instanceof Boolean){
                    resultCellType = Cell.CELL_TYPE_BOOLEAN;
                }
            }
        }
        return evaluationResult;
    }

    public int getColIndex() {
        return colIndex;
    }

    public void readCell(Cell cell){
        readCellGeneralInfo(cell);
        readCellContents(cell);
        readCellStyle(cell);
    }

    private void readCellGeneralInfo(Cell cell) {
        cellType = cell.getCellType();
        hyperlink = cell.getHyperlink();
        colIndex = cell.getColumnIndex();
        rowIndex = cell.getRowIndex();
    }

    private void readCellContents(Cell cell) {
        switch( cell.getCellType() ){
            case Cell.CELL_TYPE_STRING:
                richTextString = cell.getRichStringCellValue();
                evaluationResult = richTextString.getString();
                cellOriginalValue = richTextString.getString();
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                booleanValue = cell.getBooleanCellValue();
                evaluationResult = booleanValue;
                cellOriginalValue = booleanValue;
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if(DateUtil.isCellDateFormatted(cell)) {
                    dateValue = cell.getDateCellValue();
                    evaluationResult = dateValue;
                    cellOriginalValue = dateValue;
                } else {
                    doubleValue = cell.getNumericCellValue();
                    evaluationResult = doubleValue;
                    cellOriginalValue = doubleValue;
                }
                break;
            case Cell.CELL_TYPE_FORMULA:
                formula = cell.getCellFormula();
                evaluationResult = formula;
                cellOriginalValue = formula;
                break;
            case Cell.CELL_TYPE_ERROR:
                errorValue = cell.getErrorCellValue();
                evaluationResult = errorValue;
                cellOriginalValue = errorValue;
                break;
        }
    }

    private void readCellStyle(Cell cell) {
        style = cell.getCellStyle();
    }

    public void writeToCell(Cell cell, Context context){
        evaluate(context);
        updateCellGeneralInfo(cell);
        updateCellContents( cell );
        updateCellStyle( cell );
    }

    private void updateCellGeneralInfo(Cell cell) {
        cell.setCellType( resultCellType );
        if( hyperlink != null ){
            cell.setHyperlink( hyperlink );
        }
    }

    private void updateCellContents(Cell cell) {
        switch( resultCellType ){
            case Cell.CELL_TYPE_STRING:
                cell.setCellValue((String) evaluationResult);
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                cell.setCellValue( (Boolean)evaluationResult );
                break;
            case Cell.CELL_TYPE_NUMERIC:
                    if( evaluationResult instanceof Integer){
                        cell.setCellValue(((Integer)evaluationResult).doubleValue());
                    }else{
                        cell.setCellValue((Double) evaluationResult);
                    }
                break;
            case Cell.CELL_TYPE_FORMULA:
                cell.setCellFormula((String) evaluationResult);
                break;
            case Cell.CELL_TYPE_ERROR:
                cell.setCellErrorValue((Byte) evaluationResult);
                break;
        }
    }

    private void updateCellStyle(Cell cell) {
        cell.setCellStyle( style );
    }

    @Override
    public String toString() {
        return "CellData{" +
                "col=" + colIndex +
                ", row=" + rowIndex +
                ", source cell type=" + cellType +
                ", target cell type=" + resultCellType +
                ", cellValue=" + cellOriginalValue +
                ", eval.result=" + evaluationResult +
                '}';
    }
}
