package org.jxls.builder;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.Map;

import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.expression.ExpressionEvaluatorFactory;
import org.jxls.expression.ExpressionEvaluatorFactoryJexlImpl;
import org.jxls.formula.FastFormulaProcessor;
import org.jxls.formula.FormulaProcessor;
import org.jxls.formula.StandardFormulaProcessor;
import org.jxls.transform.JxlsTransformerFactory;

/**
 * You must call withTransformerFactory() and withTemplate().
 */
public class JxlsTemplateFillerBuilder<SELF extends JxlsTemplateFillerBuilder<SELF>> {
    private ExpressionEvaluatorFactory expressionEvaluatorFactory = new ExpressionEvaluatorFactoryJexlImpl();
    public static final String DEFAULT_EXPRESSION_BEGIN = "${";
    public static final String DEFAULT_EXPRESSION_END = "}";
    protected String expressionNotationBegin = DEFAULT_EXPRESSION_BEGIN;
    protected String expressionNotationEnd = DEFAULT_EXPRESSION_END;
    /** old name: evaluateFormulas */
    protected boolean recalculateFormulasBeforeSaving = true;
    /** old name: fullFormulaRecalculationOnOpening */
    protected boolean recalculateFormulasOnOpening = false;
    /** null: no formula processing */
    private FormulaProcessor formulaProcessor = new StandardFormulaProcessor();
    protected boolean hideTemplateSheet = false;
    protected boolean deleteTemplateSheet = false;
    private AreaBuilder areaBuilder = new XlsCommentAreaBuilder();
    protected boolean clearTemplateCells = true;
    private JxlsTransformerFactory transformerFactory;
    protected JxlsStreaming streaming = JxlsStreaming.STREAMING_OFF;
    protected InputStream template;

    public static JxlsTemplateFillerBuilder<?> newInstance() {
        return new JxlsTemplateFillerBuilder<>();
    }

    /**
     * @return all options and the template
     */
    public JxlsTemplateFiller build() {
        return new JxlsTemplateFiller(expressionEvaluatorFactory, expressionNotationBegin, expressionNotationEnd,
                recalculateFormulasBeforeSaving, recalculateFormulasOnOpening, formulaProcessor,
                hideTemplateSheet, deleteTemplateSheet, areaBuilder, clearTemplateCells, transformerFactory, streaming, template);
    }
    
    /**
     * @param data -
     * @return Excel file
     * @throws IOException 
     */
    public JxlsOutput build(Map<String, Object> data) {
        return build().fill(data);
    }
    
    public SELF withExpressionEvaluatorFactory(ExpressionEvaluatorFactory expressionEvaluatorFactory) {
        if (expressionEvaluatorFactory == null) {
            throw new IllegalArgumentException();
        }
        this.expressionEvaluatorFactory = expressionEvaluatorFactory;
        return (SELF) this;
    }
    
    public ExpressionEvaluatorFactory getExpressionEvaluatorFactory() {
        return expressionEvaluatorFactory;
    }

    public SELF withExpressionNotation(String begin, String end) {
        expressionNotationBegin = begin == null ? DEFAULT_EXPRESSION_BEGIN : begin;
        expressionNotationEnd = end == null ? DEFAULT_EXPRESSION_END : end;
        return (SELF) this;
    }

    public SELF withRecalculateFormulasBeforeSaving(boolean recalculateFormulasBeforeSaving) {
        this.recalculateFormulasBeforeSaving = recalculateFormulasBeforeSaving;
        return (SELF) this;
    }

    public SELF withRecalculateFormulasOnOpening(boolean recalculateFormulasOnOpening) {
        this.recalculateFormulasOnOpening = recalculateFormulasOnOpening;
        return (SELF) this;
    }

    public SELF withFormulaProcessor(FormulaProcessor formulaProcessor) {
        this.formulaProcessor = formulaProcessor;
        return (SELF) this;
    }
    
    public FormulaProcessor getFormulaProcessor() {
        return formulaProcessor;
    }

    public SELF withFastFormulaProcessor() {
        return withFormulaProcessor(new FastFormulaProcessor());
    }

    public SELF withHideTemplateSheet(boolean hideTemplateSheet) {
        this.hideTemplateSheet = hideTemplateSheet;
        return (SELF) this;
    }
    
    public SELF withDeleteTemplateSheet(boolean deleteTemplateSheet) {
        this.deleteTemplateSheet = deleteTemplateSheet;
        return (SELF) this;
    }

    public SELF withAreaBuilder(AreaBuilder areaBuilder) {
        this.areaBuilder = areaBuilder;
        return (SELF) this;
    }

    public AreaBuilder getAreaBuilder() {
        return areaBuilder;
    }
    
    public SELF withClearTemplateCells(boolean clearTemplateCells) {
        this.clearTemplateCells = clearTemplateCells;
        return (SELF) this;
    }
    
    public SELF withTransformerFactory(JxlsTransformerFactory transformerFactory) {
        this.transformerFactory = transformerFactory;
        return (SELF) this;
    }

    public JxlsTransformerFactory getTransformerFactory() {
        return transformerFactory;
    }

    public SELF withStreaming(JxlsStreaming streaming) {
        this.streaming = streaming;
        return (SELF) this;
    }

    public SELF withTemplate(InputStream template) {
        this.template = template;
        return (SELF) this;
    }

    public SELF withTemplate(URL template) {
        try {
            return withTemplate(template.openStream());
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public SELF withTemplate(File template) {
        try {
            return withTemplate(new FileInputStream(template));
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
    }

    public SELF withTemplate(String templateFileName) {
        return withTemplate(new File(templateFileName));
    }
}
