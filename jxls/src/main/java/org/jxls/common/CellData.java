package org.jxls.common;

import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.ZonedDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.jxls.area.XlsArea;
import org.jxls.builder.xls.JxlsCommentException;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.expression.ExpressionEvaluator;
import org.jxls.formula.AbstractFormulaProcessor;
import org.jxls.transform.TransformationConfig;
import org.jxls.transform.Transformer;

/**
 * Represents an Excel cell data holder and cell value evaluator
 * 
 * @author Leonid Vysochyn
 */
public class CellData {
    private static final String USER_FORMULA_PREFIX = "$[";
    private static final String USER_FORMULA_SUFFIX = "]";
    private static final String ATTR_PREFIX = "(";
    private static final String ATTR_SUFFIX = ")";
    public static final String JX_PARAMS_PREFIX = XlsCommentAreaBuilder.COMMAND_PREFIX + "params";
    /*
     * In addition to normal (straight) single and double quotes, this regex
     * includes the following commonly occurring quote-like characters (some
     * of which have been observed in recent versions of LibreOffice):
     *
     * U+201C - LEFT DOUBLE QUOTATION MARK
     * U+201D - RIGHT DOUBLE QUOTATION MARK
     * U+201E - DOUBLE LOW-9 QUOTATION MARK
     * U+201F - DOUBLE HIGH-REVERSED-9 QUOTATION MARK
     * U+2033 - DOUBLE PRIME
     * U+2036 - REVERSED DOUBLE PRIME
     * U+2018 - LEFT SINGLE QUOTATION MARK
     * U+2019 - RIGHT SINGLE QUOTATION MARK
     * U+201A - SINGLE LOW-9 QUOTATION MARK
     * U+201B - SINGLE HIGH-REVERSED-9 QUOTATION MARK
     * U+2032 - PRIME
     * U+2035 - REVERSED PRIME
     */
    private static final String ATTR_REGEX = "\\s*\\w+\\s*=\\s*([\"|'\u201C\u201D\u201E\u201F\u2033\u2036\u2018\u2019\u201A\u201B\u2032\u2035])(?:(?!\\1).)*\\1";
    private static final Pattern ATTR_REGEX_PATTERN = Pattern.compile(ATTR_REGEX);
    private static final String FORMULA_STRATEGY_PARAM = "formulaStrategy";
    private static final String DEFAULT_VALUE = "defaultValue";

    public enum CellType {
        STRING, NUMBER, BOOLEAN, DATE, LOCAL_DATE, LOCAL_TIME, LOCAL_DATETIME, ZONED_DATETIME, INSTANT, FORMULA, BLANK, ERROR
    }

    public enum FormulaStrategy {
        DEFAULT, BY_COLUMN, BY_ROW
    }

    private Map<String, String> attrMap;
    protected CellRef cellRef;
    protected Object cellValue;
    protected CellType cellType;
    private String cellComment;
    protected String formula;
    protected Object evaluationResult;
    protected CellType targetCellType;
    private FormulaStrategy formulaStrategy = FormulaStrategy.DEFAULT;
    private String defaultValue;
    protected XlsArea area;
    private List<CellRef> targetPos = new ArrayList<CellRef>();
    private List<AreaRef> targetParentAreaRef = new ArrayList<>();
    private Transformer transformer;

    private List<String> evaluatedFormulas = new ArrayList<>();

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

    protected void updateFormulaValue() {
        if (cellType == CellType.FORMULA) {
            formula = cellValue != null ? cellValue.toString() : "";
        } else if (cellType == CellType.STRING && cellValue != null && isUserFormula(cellValue.toString())) {
            formula = cellValue.toString().substring(2, cellValue.toString().length() - 1);
        }
    }

    public Transformer getTransformer() {
        return transformer;
    }

    public void setTransformer(Transformer transformer) {
        this.transformer = transformer;
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

    public void setEvaluationResult(Object evaluationResult) {
        this.evaluationResult = evaluationResult;
    }

    private ExpressionEvaluator getExpressionEvaluator() {
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

    public String getCellComment() {
        return cellComment;
    }

    public void setCellComment(String cellComment) {
        this.cellComment = cellComment;
    }

    protected boolean isJxlsParamsComment(String cellComment) {
        return cellComment.trim().startsWith(JX_PARAMS_PREFIX);
    }

    public String getSheetName() {
        return cellRef.getSheetName();
    }

    public CellRef getCellRef() {
        return cellRef;
    }

    public CellType getCellType() {
        return cellType;
    }

    public void setCellType(CellType cellType) {
        this.cellType = cellType;
    }

    public Object getCellValue() {
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

    public List<String> getEvaluatedFormulas() {
        return evaluatedFormulas;
    }

    public boolean isFormulaCell() {
        return formula != null;
    }

    public boolean isParameterizedFormulaCell() {
        return isFormulaCell() && isUserFormula(cellValue.toString());
    }

    public boolean isJointedFormulaCell() {
        return isParameterizedFormulaCell() && formulaContainsJointedCellRef(cellValue.toString());
    }
    
    /**
     * Checks if the formula contains jointed cell references
     * Jointed references have format U_(cell1, cell2) e.g. $[SUM(U_(F8,F13))]
     * @param formula string
     * @return true if the formula contains jointed cell references
     */
    protected boolean formulaContainsJointedCellRef(String formula) {
        return AbstractFormulaProcessor.regexJointedCellRefPattern.matcher(formula).find();
    }

    public boolean addTargetPos(CellRef cellRef) {
        return targetPos.add(cellRef);
    }

    public void addTargetParentAreaRef(AreaRef areaRef) {
        targetParentAreaRef.add(areaRef);
    }

    public List<AreaRef> getTargetParentAreaRef() {
        return targetParentAreaRef;
    }

    public void setEvaluatedFormulas(List<String> evaluatedFormulas) {
        this.evaluatedFormulas = evaluatedFormulas;
    }

    /**
     * @return a list of cell refs into which the current cell was transformed
     */
    public List<CellRef> getTargetPos() {
        return targetPos;
    }

    public void resetTargetPos() {
        targetPos.clear();
        targetParentAreaRef.clear();
    }

    public Object evaluate(Context context) {
        targetCellType = cellType;
        if (cellType == CellType.STRING && cellValue != null) {
            String strValue = cellValue.toString();
            if (isUserFormula(strValue)) {
                String formulaStr = strValue.substring(USER_FORMULA_PREFIX.length(), strValue.length() - USER_FORMULA_SUFFIX.length());
                evaluate(formulaStr, context);
                if (evaluationResult != null) {
                    targetCellType = CellType.FORMULA;
                    formula = evaluationResult.toString();
                    evaluatedFormulas.add(formula);
                }
            } else {
                evaluate(strValue, context);
            }
            if (evaluationResult == null) {
                targetCellType = CellType.BLANK;
            }
        }
        return evaluationResult;
    }

    private static boolean isUserFormula(String str) {
        return str.startsWith(USER_FORMULA_PREFIX) && str.endsWith(USER_FORMULA_SUFFIX);
    }
    
    private void evaluate(String strValue, Context context) {
        StringBuffer sb = new StringBuffer(); // TODO StringBuilder
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
        while (exprMatcher.find()) {
            endOffset = exprMatcher.end();
            matchCount++;
            matchedString = exprMatcher.group();
            expression = matchedString.substring(beginExpressionLength, matchedString.length() - endExpressionLength);
            lastMatchEvalResult = evaluator.evaluate(expression, context.toMap());
            exprMatcher.appendReplacement(sb,
                    Matcher.quoteReplacement(lastMatchEvalResult != null ? lastMatchEvalResult.toString() : ""));
        }
        String lastStringResult = lastMatchEvalResult != null ? lastMatchEvalResult.toString() : "";
        boolean isAppendTail = matchCount == 1 && endOffset < strValue.length();
        if (matchCount > 1 || isAppendTail) {
            exprMatcher.appendTail(sb);
            evaluationResult = sb.toString();
        } else if (matchCount == 1) {
            if (sb.length() > lastStringResult.length()) {
                evaluationResult = sb.toString();
            } else {
                evaluationResult = lastMatchEvalResult;
                setTargetCellType();
            }
        } else if (matchCount == 0) {
            evaluationResult = strValue;
        }
    }

    private void setTargetCellType() {
        if (evaluationResult instanceof Number) {
            targetCellType = CellType.NUMBER;
        } else if (evaluationResult instanceof Boolean) {
            targetCellType = CellType.BOOLEAN;
        } else if (evaluationResult instanceof Date) {
            targetCellType = CellType.DATE;
        } else if (evaluationResult instanceof LocalDate) {
            targetCellType = CellType.LOCAL_DATE;
        } else if (evaluationResult instanceof LocalTime) {
            targetCellType = CellType.LOCAL_TIME;
        } else if (evaluationResult instanceof LocalDateTime) {
            targetCellType = CellType.LOCAL_DATETIME;
        } else if (evaluationResult instanceof ZonedDateTime) {
            targetCellType = CellType.ZONED_DATETIME;
        } else if (evaluationResult instanceof Instant) {
            targetCellType = CellType.INSTANT;
        }
    }

    /**
     * The method parses jx:params attribute from a cell comment
     * <p>jx:params can be used e.g.</p><ul>
     * <li>to set {@link FormulaStrategy} via 'formulaStrategy' param</li>
     * <li>to set the formula default value via 'defaultValue' param</li></ul>
     * 
     * @param cellComment the comment string
     */
    protected void processJxlsParams(String cellComment) {
        int nameEndIndex = cellComment.indexOf(ATTR_PREFIX, JX_PARAMS_PREFIX.length());
        if (nameEndIndex < 0) {
            throw new JxlsCommentException("Failed to parse jxls params '" + cellComment + "' at " + cellRef.getCellName()
                    + ". Expected '" + ATTR_PREFIX + "' symbol.");
        }
        attrMap = buildAttrMap(cellComment, nameEndIndex);
        if (attrMap.containsKey(FORMULA_STRATEGY_PARAM)) {
            initFormulaStrategy(attrMap.get(FORMULA_STRATEGY_PARAM));
        }
        if (attrMap.containsKey(DEFAULT_VALUE)) {
            defaultValue = attrMap.get(DEFAULT_VALUE);
        }
    }

    private Map<String, String> buildAttrMap(String paramsLine, int nameEndIndex) {
        int paramsEndIndex = paramsLine.lastIndexOf(ATTR_SUFFIX);
        if (paramsEndIndex < 0) {
            throw new JxlsCommentException("Failed to parse params line '" + paramsLine + "' at " + cellRef.getCellName()
                    + ". Expected '" + ATTR_SUFFIX + "' symbol.");
        }
        String attrString = paramsLine.substring(nameEndIndex + 1, paramsEndIndex).trim();
        return parseCommandAttributes(attrString);
    }

    private Map<String, String> parseCommandAttributes(String attrString) {
        Map<String, String> attrMap = new LinkedHashMap<String, String>();
        Matcher attrMatcher = ATTR_REGEX_PATTERN.matcher(attrString);
        while (attrMatcher.find()) {
            String attrData = attrMatcher.group();
            int attrNameEndIndex = attrData.indexOf("=");
            String attrName = attrData.substring(0, attrNameEndIndex).trim();
            String attrValuePart = attrData.substring(attrNameEndIndex + 1).trim();
            String attrValue = attrValuePart.substring(1, attrValuePart.length() - 1);
            attrMap.put(attrName, attrValue);
        }
        return attrMap;
    }

    private void initFormulaStrategy(String formulaStrategyValue) {
        try {
            this.formulaStrategy = FormulaStrategy.valueOf(formulaStrategyValue);
        } catch (IllegalArgumentException e) {
            throw new JxlsException("Cannot parse formula strategy value at " + cellRef.getCellName(), e);
        }
    }

    @Override
    public String toString() {
        return "CellData{" +
                cellRef +
                ", cellType=" + cellType +
                ", cellValue=" + cellValue +
                '}';
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) {
            return true;
        }
        if (!(o instanceof CellData)) {
            return false;
        }

        CellData cellData = (CellData) o;
        if (cellType != cellData.cellType) {
            return false;
        }
        if (cellValue != null ? !cellValue.equals(cellData.cellValue) : cellData.cellValue != null) {
            return false;
        }
        return cellRef != null ? cellRef.equals(cellData.cellRef) : cellData.cellRef == null;
    }

    @Override
    public int hashCode() {
        int result = cellRef != null ? cellRef.hashCode() : 0;
        result = 31 * result + (cellValue != null ? cellValue.hashCode() : 0);
        result = 31 * result + (cellType != null ? cellType.hashCode() : 0);
        return result;
    }
}
