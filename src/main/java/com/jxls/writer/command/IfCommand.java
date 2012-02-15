package com.jxls.writer.command;

import com.jxls.writer.common.CellRef;
import com.jxls.writer.common.Size;
import com.jxls.writer.expression.ExpressionEvaluator;
import com.jxls.writer.expression.JexlExpressionEvaluator;

/**
 * Date: Sep 11, 2009
 *
 * @author Leonid Vysochyn
 */
public class IfCommand extends AbstractCommand {

    String condition;
    Boolean conditionResult;
    Area ifArea;
    Area elseArea;

    public IfCommand(String condition) {
        this.condition = condition;
    }

    public IfCommand(String condition, Area ifArea, Area elseArea){
        this.condition = condition;
        this.ifArea = ifArea != null ? ifArea : XlsArea.EMPTY_AREA;
        this.elseArea = elseArea != null ? elseArea : XlsArea.EMPTY_AREA;
        super.addArea(this.ifArea);
        super.addArea(this.elseArea);
    }

    public IfCommand(String condition, XlsArea ifArea) {
        this(condition, ifArea, XlsArea.EMPTY_AREA);
    }

    public String getName() {
        return "if";
    }

    public String getCondition() {
        return condition;
    }

    @Override
    public void addArea(Area area) {
        if( areaList.size() >= 2 ){
            throw new IllegalArgumentException("Cannot add any more areas to this IfCommand. You can add only 1 area for 'if' part and 1 area for 'else' part");
        }
        if(areaList.isEmpty()){
            ifArea = area;
        }else {
            elseArea = area;
        }
        super.addArea(area);
    }

    public Size applyAt(CellRef cellRef, Context context) {
        conditionResult = isConditionTrue(context);
        if( conditionResult ){
            return ifArea.applyAt(cellRef, context);
        }else{
            return elseArea.applyAt(cellRef, context);
        }
    }

    public Boolean isConditionTrue(Context context){
        ExpressionEvaluator expressionEvaluator = new JexlExpressionEvaluator(context.toMap());
        Object conditionResult = expressionEvaluator.evaluate(condition);
        if( !(conditionResult instanceof Boolean) ){
            throw new RuntimeException("If command condition is not a boolean value - " + condition);
        }
        return (Boolean)conditionResult;
    }
}
