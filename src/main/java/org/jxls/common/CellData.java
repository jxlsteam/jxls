package org.jxls.common;

import org.jxls.area.XlsArea;
import org.jxls.expression.ExpressionEvaluator;
import org.jxls.transform.TransformationConfig;
import org.jxls.transform.Transformer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Represents an excel cell data holder and cell value evaluator
 * @author Leonid Vysochyn
 *         Date: 2/3/12
 */
public class CellData {
    public static final String USER_FORMULA_PREFIX = "$[";
    public static final String USER_FORMULA_SUFFIX = "]";
    private static final String ATTR_PREFIX = "(";
    private static final String ATTR_SUFFIX = ")";
    public static final String JX_PARAMS_PREFIX = "jx:params";
    private static final String ATTR_REGEX = "\\s*\\w+\\s*=\\s*([\"|'])(?:(?!\\1).)*\\1";
    private static final Pattern ATTR_REGEX_PATTERN = Pattern.compile(ATTR_REGEX);
    public static final String FORMULA_STRATEGY_PARAM = "formulaStrategy";
    public static final String DEFAULT_VALUE = "defaultValue";
    private Map<String, String> attrMap;

    public enum CellType {
        STRING, NUMBER, BOOLEAN, DATE, FORMULA, BLANK, ERROR
    }

    public enum FormulaStrategy {
        DEFAULT, BY_COLUMN, BY_ROW
    }

    static Logger logger = LoggerFactory.getLogger(CellData.class);

    protected CellRef cellRef;
    protected Object cellValue;
    protected CellType cellType;
    protected String cellComment;

    protected String formula;
    protected Object evaluationResult;
    protected CellType targetCellType;
    protected FormulaStrategy formulaStrategy = FormulaStrategy.DEFAULT;
    protected String defaultValue;

    protected XlsArea area;

    List<CellRef> targetPos = new ArrayList<CellRef> ();
    List<AreaRef> targetParentAreaRef = new ArrayList<>();

    Transformer transformer;

    public Transformer getTransformer() {
        return transformer;
    }

    public void setTransformer(Transformer transformer) {
        this.transformer = transformer;
    }

    public CellData(CellRef cellRef) {
        this.cellRef = cellRef;
    }

    public CellData(String sheetName, int row, int col, CellType cellType, Object cellValue) {
        this.cellRef = new CellRef(sheetName, row, col);
        this.cellType = cellType;
        this.cellValue = cellValue;
        updateFormulaValue();
    }

    public CellData(CellRef cellRef, CellType cellType, Object cellValue) {
        this.cellRef = cellRef;
        this.cellType = cellType;
        this.cellValue = cellValue;
        updateFormulaValue();
    }

    public CellData(String sheetName, int row, int col) {
        this(sheetName, row, col, CellType.BLANK, null);
    }

    public XlsArea getArea() {
        return area;
    }

    public void setArea(XlsArea area) {
        this.area = area;
    }

    public Map<String, String> getAttrMap() {
        return attrMap;
    }

    public void setAttrMap(Map<String, String> attrMap) {
        this.attrMap = attrMap;
    }

    public Object evaluate(Context context){
        targetCellType = cellType;
        if( cellType == CellType.STRING && cellValue != null){
            String strValue = cellValue.toString();
            if( isUserFormula(strValue) ){
                String formulaStr = strValue.substring(2, strValue.length()-1);
                evaluate(formulaStr, context);
                if( evaluationResult != null ){
                    targetCellType = CellType.FORMULA;
                    formula = evaluationResult.toString();
                }
            }else{
                evaluate(strValue, context);
            }
            if(evaluationResult == null){
                targetCellType = CellType.BLANK;
            }
        }
        return evaluationResult;
    }

    private ExpressionEvaluator getExpressionEvaluator(){
        return transformer.getTransformationConfig().getExpressionEvaluator();
    }

    public FormulaStrategy getFormulaStrategy() {
        return formulaStrategy;
    }

    public void setFormulaStrategy(FormulaStrategy formulaStrategy) {
        this.formulaStrategy = formulaStrategy;
    }

    public String getDefaultValue() {
        return defaultValue;
    }

    public void setDefaultValue(String defaultValue) {
        this.defaultValue = defaultValue;
    }

    void evaluate(String strValue, Context context) {
        StringBuffer sb = new StringBuffer();
        TransformationConfig transformationConfig = transformer.getTransformationConfig();
        int beginExpressionLength = transformationConfig.getExpressionNotationBegin().length();
        int endExpressionLength = transformationConfig.getExpressionNotationEnd().length();
        Matcher exprMatcher = transformationConfig.getExpressionNotationPattern().matcher(strValue);
        ExpressionEvaluator evaluator = getExpressionEvaluator();
        String matchedString;
        String expression;
        Object lastMatchEvalResult = null;
        int matchCount = 0;
        int endOffset = 0;
        while(exprMatcher.find()){
            endOffset = exprMatcher.end();
            matchCount++;
            matchedString = exprMatcher.group();
            expression = matchedString.substring(beginExpressionLength, matchedString.length() - endExpressionLength);
            lastMatchEvalResult = evaluator.evaluate(expression, context.toMap());
            exprMatcher.appendReplacement(sb, Matcher.quoteReplacement( lastMatchEvalResult != null ? lastMatchEvalResult.toString() : "" ));
        }
        String lastStringResult = lastMatchEvalResult != null ? lastMatchEvalResult.toString() : "";
        boolean isAppendTail = matchCount == 1 && endOffset < strValue.length();
        if( matchCount > 1 || isAppendTail){
            exprMatcher.appendTail(sb);
            evaluationResult = sb.toString();
        }else if(matchCount == 1){
            if(sb.length() > lastStringResult.length()){
                evaluationResult = sb.toString();
            }else {
                evaluationResult = lastMatchEvalResult;
                if (evaluationResult instanceof Number) {
                    targetCellType = CellType.NUMBER;
                } else if (evaluationResult instanceof Boolean) {
                    targetCellType = CellType.BOOLEAN;
                } else if (evaluationResult instanceof Date) {
                    targetCellType = CellType.DATE;
                }
            }
        }else if(matchCount == 0){
            evaluationResult = strValue;
        }
    }

    public String getCellComment() {
        return cellComment;
    }

    public void setCellComment(String cellComment) {
        this.cellComment = cellComment;
    }

    public boolean isJxlsParamsComment(String cellComment) {
        return cellComment.startsWith(JX_PARAMS_PREFIX);
    }

    public void processJxlsParams(String cellComment) {
        int nameEndIndex = cellComment.indexOf(ATTR_PREFIX, JX_PARAMS_PREFIX.length());
        if (nameEndIndex < 0) {
            String errMsg = "Failed to parse jxls params [" + cellComment + "] at " + cellRef.getCellName() +
                    ". Expected '" + ATTR_PREFIX + "' symbol.";
            logger.error(errMsg);
            throw new IllegalStateException(errMsg);
        }
        attrMap = buildAttrMap(cellComment, nameEndIndex);
        if( attrMap.containsKey(FORMULA_STRATEGY_PARAM) ){
            initFormulaStrategy(attrMap.get(FORMULA_STRATEGY_PARAM));
        }
        if( attrMap.containsKey(DEFAULT_VALUE) ){
            defaultValue = attrMap.get(DEFAULT_VALUE);
        }
    }

    private void initFormulaStrategy(String formulaStrategyValue) {
        try {
            this.formulaStrategy = FormulaStrategy.valueOf(formulaStrategyValue);
        } catch (IllegalArgumentException e) {
            throw new JxlsException("Cannot parse formula strategy value at " + cellRef.getCellName(), e);
        }
    }

    private Map<String, String> buildAttrMap(String paramsLine, int nameEndIndex) {
        int paramsEndIndex = paramsLine.lastIndexOf(ATTR_SUFFIX);
        if(paramsEndIndex < 0 ){
            String errMsg = "Failed to parse params line [" + paramsLine + "] at " + cellRef.getCellName() +
                    ". Expected '" + ATTR_SUFFIX + "' symbol.";
            logger.error(errMsg);
            throw new IllegalArgumentException(errMsg);
        }
        String attrString = paramsLine.substring(nameEndIndex + 1, paramsEndIndex).trim();
        return parseCommandAttributes(attrString);
    }

    private Map<String, String> parseCommandAttributes(String attrString) {
        Map<String,String> attrMap = new LinkedHashMap<String, String>();
        Matcher attrMatcher = ATTR_REGEX_PATTERN.matcher(attrString);
        while(attrMatcher.find()){
            String attrData = attrMatcher.group();
            int attrNameEndIndex = attrData.indexOf("=");
            String attrName = attrData.substring(0, attrNameEndIndex).trim();
            String attrValuePart = attrData.substring(attrNameEndIndex + 1).trim();
            String attrValue = attrValuePart.substring(1, attrValuePart.length() - 1);
            attrMap.put(attrName, attrValue);
        }
        return attrMap;
    }

    public String getSheetName() {
        return cellRef.getSheetName();
    }

    protected void updateFormulaValue() {
        if( cellType == CellType.FORMULA ){
            formula = cellValue != null ? cellValue.toString() : "";
        }else if( cellType == CellType.STRING && cellValue != null && isUserFormula(cellValue.toString())){
            formula = cellValue.toString().substring(2, cellValue.toString().length() - 1);
        }
    }

    public CellRef getCellRef(){
        return cellRef;
    }

    public CellType getCellType() {
        return cellType;
    }

    public void setCellType(CellType cellType) {
        this.cellType = cellType;
    }

    public Object getCellValue(){
        return cellValue;
    }

    public int getRow() {
        return cellRef.getRow();
    }

    public int getCol() {
        return cellRef.getCol();
    }

    public String getFormula() {
        return formula;
    }

    public void setFormula(String formula) {
        this.formula = formula;
    }

    public boolean isFormulaCell(){
        return formula != null;
    }

    public static boolean isUserFormula(String str) {
        return str.startsWith(CellData.USER_FORMULA_PREFIX) && str.endsWith(CellData.USER_FORMULA_SUFFIX);
    }

    public boolean addTargetPos(CellRef cellRef){
        return targetPos.add(cellRef);
    }

    public void addTargetParentAreaRef(AreaRef areaRef){
        targetParentAreaRef.add(areaRef);
    }

    public List<AreaRef> getTargetParentAreaRef() {
        return targetParentAreaRef;
    }

    public List<CellRef> getTargetPos(){
        return targetPos;
    }



    public void resetTargetPos(){
        targetPos.clear();
        targetParentAreaRef.clear();
    }


    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (!(o instanceof CellData)) return false;

        CellData cellData = (CellData) o;

        if (cellType != cellData.cellType) return false;
        if (cellValue != null ? !cellValue.equals(cellData.cellValue) : cellData.cellValue != null) return false;
        if (cellRef != null ? !cellRef.equals(cellData.cellRef) : cellData.cellRef != null) return false;

        return true;
    }

    @Override
    public int hashCode() {
        int result = cellRef != null ? cellRef.hashCode() : 0;
        result = 31 * result + (cellValue != null ? cellValue.hashCode() : 0);
        result = 31 * result + (cellType != null ? cellType.hashCode() : 0);
        return result;
    }

    @Override
    public String toString() {
        return "CellData{" +
                cellRef +
                ", cellType=" + cellType +
                ", cellValue=" + cellValue +
                '}';
    }
}
