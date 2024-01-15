package org.jxls.builder;

import java.util.List;
import java.util.Map;

import org.jxls.command.Command;
import org.jxls.common.NeedsPublicContext;
import org.jxls.common.RunVarAccess;
import org.jxls.expression.ExpressionEvaluatorFactory;
import org.jxls.formula.FormulaProcessor;
import org.jxls.logging.JxlsLogger;
import org.jxls.transform.JxlsTransformerFactory;
import org.jxls.transform.PreWriteAction;

/**
 * Internal transport object for delivering the builder options to the template filler
 */
public class JxlsOptions {
    private final ExpressionEvaluatorFactory expressionEvaluatorFactory;
    private final String expressionNotationBegin;
    private final String expressionNotationEnd;
    private final JxlsLogger logger;
    private final FormulaProcessor formulaProcessor;
    private final boolean updateCellDataArea;
    private final boolean ignoreColumnProps;
    private final boolean ignoreRowProps;
    private final boolean recalculateFormulasBeforeSaving;
    private final boolean recalculateFormulasOnOpening;
    private final KeepTemplateSheet keepTemplateSheet;
    private final AreaBuilder areaBuilder;
    private final Map<String, Class<? extends Command>> commands;
    private final boolean clearTemplateCells;
    private final JxlsTransformerFactory transformerFactory;
    private final JxlsStreaming streaming;
    private final List<NeedsPublicContext> needsContextList;
    private final List<PreWriteAction> preWriteActions;
    private final RunVarAccess runVarAccess;

    public JxlsOptions(ExpressionEvaluatorFactory expressionEvaluatorFactory, String expressionNotationBegin,
            String expressionNotationEnd, JxlsLogger logger, FormulaProcessor formulaProcessor, boolean updateCellDataArea,
            boolean ignoreColumnProps, boolean ignoreRowProps, boolean recalculateFormulasBeforeSaving,
            boolean recalculateFormulasOnOpening, KeepTemplateSheet keepTemplateSheet, AreaBuilder areaBuilder,
            Map<String, Class<? extends Command>> commands, boolean clearTemplateCells,
            JxlsTransformerFactory transformerFactory, JxlsStreaming streaming, List<NeedsPublicContext> needsContextList,
            List<PreWriteAction> preWriteActions, RunVarAccess runVarAccess) {
        this.expressionEvaluatorFactory = expressionEvaluatorFactory;
        this.expressionNotationBegin = expressionNotationBegin;
        this.expressionNotationEnd = expressionNotationEnd;
        this.logger = logger;
        this.formulaProcessor = formulaProcessor;
        this.updateCellDataArea = updateCellDataArea;
        this.ignoreColumnProps = ignoreColumnProps;
        this.ignoreRowProps = ignoreRowProps;
        this.recalculateFormulasBeforeSaving = recalculateFormulasBeforeSaving;
        this.recalculateFormulasOnOpening = recalculateFormulasOnOpening;
        this.keepTemplateSheet = keepTemplateSheet;
        this.areaBuilder = areaBuilder;
        this.commands = commands;
        this.clearTemplateCells = clearTemplateCells;
        this.transformerFactory = transformerFactory;
        this.streaming = streaming;
        this.needsContextList = needsContextList;
        this.preWriteActions = preWriteActions;
        this.runVarAccess = runVarAccess;
    }

    public ExpressionEvaluatorFactory getExpressionEvaluatorFactory() {
        return expressionEvaluatorFactory;
    }

    public String getExpressionNotationBegin() {
        return expressionNotationBegin;
    }

    public String getExpressionNotationEnd() {
        return expressionNotationEnd;
    }

    public JxlsLogger getLogger() {
        return logger;
    }

    public FormulaProcessor getFormulaProcessor() {
        return formulaProcessor;
    }

    public boolean isUpdateCellDataArea() {
        return updateCellDataArea;
    }

    public boolean isIgnoreColumnProps() {
        return ignoreColumnProps;
    }

    public boolean isIgnoreRowProps() {
        return ignoreRowProps;
    }

    public boolean isRecalculateFormulasBeforeSaving() {
        return recalculateFormulasBeforeSaving;
    }

    public boolean isRecalculateFormulasOnOpening() {
        return recalculateFormulasOnOpening;
    }

    public KeepTemplateSheet getKeepTemplateSheet() {
        return keepTemplateSheet;
    }

    public AreaBuilder getAreaBuilder() {
        return areaBuilder;
    }

    public Map<String, Class<? extends Command>> getCommands() {
        return commands;
    }

    public boolean isClearTemplateCells() {
        return clearTemplateCells;
    }

    public JxlsTransformerFactory getTransformerFactory() {
        return transformerFactory;
    }

    public JxlsStreaming getStreaming() {
        return streaming;
    }

    public List<NeedsPublicContext> getNeedsPublicContextList() {
        return needsContextList;
    }

    public List<PreWriteAction> getPreWriteActions() {
        return preWriteActions;
    }

    public RunVarAccess getRunVarAccess() {
        return runVarAccess;
    }
}
