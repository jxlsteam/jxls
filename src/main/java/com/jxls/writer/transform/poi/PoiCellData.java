package com.jxls.writer.transform.poi;

import com.jxls.writer.CellData;
import com.jxls.writer.command.Context;
import com.jxls.writer.expression.ExpressionEvaluator;
import com.jxls.writer.expression.JexlExpressionEvaluator;
import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Date;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author Leonid Vysochyn
 *         Date: 1/23/12 6:39 PM
 */
public class PoiCellData extends CellData {
    static Logger logger = LoggerFactory.getLogger(PoiCellData.class);

    protected static final String REGEX_EXPRESSION = "\\$\\{[^}]*}";
    protected static final Pattern REGEX_EXPRESSION_PATTERN = Pattern.compile(REGEX_EXPRESSION);


    RichTextString richTextString;
    boolean booleanValue;

    int poiCellType;
    int resultCellType;
    private Date dateValue;
    private double doubleValue;
    private CellStyle style;
    private Hyperlink hyperlink;
    private byte errorValue;


    public static PoiCellData createCellData(Cell cell){
        PoiCellData cellData = new PoiCellData();
        cellData.readCell(cell);
        return cellData;
    }

    public boolean isFormulaCell(){
        if(poiCellType == Cell.CELL_TYPE_FORMULA ) return true;
        return richTextString != null && isUserFormula(richTextString.getString());
    }
    
    public Object evaluate(Context context){
        resultCellType = poiCellType;
        if( richTextString != null){
            String strValue = richTextString.getString();
            if( isUserFormula(strValue) ){
                String formulaStr = strValue.substring(2, strValue.length()-1);
                evaluate(formulaStr, context);
                if( evaluationResult != null ){
                    evaluationResult = evaluationResult.toString();
                    resultCellType = Cell.CELL_TYPE_FORMULA;
                }
            }else{
                evaluate(strValue, context);
            }
            if(evaluationResult == null){
                resultCellType = Cell.CELL_TYPE_BLANK;
            }
        }
        return evaluationResult;
    }

    void evaluate(String strValue, Context context) {
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
        }else if(matchCount == 0){
            evaluationResult = strValue;
        }
    }

    public void readCell(Cell cell){
        readCellGeneralInfo(cell);
        readCellContents(cell);
        readCellStyle(cell);
    }

    private void readCellGeneralInfo(Cell cell) {
        poiCellType = cell.getCellType();
        hyperlink = cell.getHyperlink();
        col = cell.getColumnIndex();
        row = cell.getRowIndex();
    }

    private void readCellContents(Cell cell) {
        switch( cell.getCellType() ){
            case Cell.CELL_TYPE_STRING:
                richTextString = cell.getRichStringCellValue();
                evaluationResult = richTextString.getString();
                cellValue = richTextString.getString();
                cellType = CellType.STRING;
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                booleanValue = cell.getBooleanCellValue();
                evaluationResult = booleanValue;
                cellValue = booleanValue;
                cellType = CellType.BOOLEAN;
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if(DateUtil.isCellDateFormatted(cell)) {
                    dateValue = cell.getDateCellValue();
                    evaluationResult = dateValue;
                    cellValue = dateValue;
                    cellType = CellType.DATE;
                } else {
                    doubleValue = cell.getNumericCellValue();
                    evaluationResult = doubleValue;
                    cellValue = doubleValue;
                    cellType = CellType.NUMBER;
                }
                break;
            case Cell.CELL_TYPE_FORMULA:
                formula = cell.getCellFormula();
                evaluationResult = formula;
                cellValue = formula;
                cellType = CellType.FORMULA;
                break;
            case Cell.CELL_TYPE_ERROR:
                errorValue = cell.getErrorCellValue();
                evaluationResult = errorValue;
                cellValue = errorValue;
                cellType = CellType.ERROR;
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
                try{
                    cell.setCellFormula((String) evaluationResult);
                }catch(FormulaParseException e){
                    String formulaString = "";
                    try{
                        formulaString = evaluationResult.toString();
                        logger.error("Failed to set cell formula " + formulaString + " for cell " + this.toString(), e);
                        cell.setCellType(Cell.CELL_TYPE_STRING);
                        cell.setCellValue(formulaString);
                    }catch(Exception ex){
                        logger.warn("Failed to convert formula to string for cell " + this.toString());
                    }
                }
                break;
            case Cell.CELL_TYPE_ERROR:
                cell.setCellErrorValue((Byte) evaluationResult);
                break;
        }
    }

    private void updateCellStyle(Cell cell) {
        cell.setCellStyle( style );
    }

}
