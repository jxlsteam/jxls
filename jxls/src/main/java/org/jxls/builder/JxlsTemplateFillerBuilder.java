package org.jxls.builder;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.HashMap;
import java.util.Map;

import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.command.Command;
import org.jxls.common.ExceptionHandler;
import org.jxls.common.JxlsException;
import org.jxls.expression.ExpressionEvaluatorFactory;
import org.jxls.expression.ExpressionEvaluatorFactoryJexlImpl;
import org.jxls.formula.FastFormulaProcessor;
import org.jxls.formula.FormulaProcessor;
import org.jxls.formula.StandardFormulaProcessor;
import org.jxls.transform.JxlsTransformerFactory;
import org.jxls.util.CannotOpenWorkbookException;

/**
 * You must call withTransformerFactory() and withTemplate().
 */
public class JxlsTemplateFillerBuilder<SELF extends JxlsTemplateFillerBuilder<SELF>> {
    private ExpressionEvaluatorFactory expressionEvaluatorFactory = new ExpressionEvaluatorFactoryJexlImpl();
    public static final String DEFAULT_EXPRESSION_BEGIN = "${";
    public static final String DEFAULT_EXPRESSION_END = "}";
    protected String expressionNotationBegin = DEFAULT_EXPRESSION_BEGIN;
    protected String expressionNotationEnd = DEFAULT_EXPRESSION_END;
    protected ExceptionHandler exceptionHandler;
    /** null: no formula processing */
    private FormulaProcessor formulaProcessor = new StandardFormulaProcessor();
    protected boolean ignoreColumnProps = false;
    protected boolean ignoreRowProps = false;
    /** old name: evaluateFormulas */
    protected boolean recalculateFormulasBeforeSaving = true;
    /** old name: fullFormulaRecalculationOnOpening */
    protected boolean recalculateFormulasOnOpening = false;
    /** old names: deleteTemplateSheet, hideTemplateSheet */
    protected KeepTemplateSheet keepTemplateSheet = KeepTemplateSheet.DELETE;
    private AreaBuilder areaBuilder = new XlsCommentAreaBuilder();
    protected final Map<String, Class<? extends Command>> commands = new HashMap<>();
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
    	if (transformerFactory == null) {
    		throw new JxlsException("Please call withTransformerFactory()");
    	} else if (template == null) {
    		throw new JxlsException("Please call withTemplate()");
    	}
        return new JxlsTemplateFiller(expressionEvaluatorFactory, expressionNotationBegin, expressionNotationEnd,
        		exceptionHandler, formulaProcessor, ignoreColumnProps, ignoreRowProps, recalculateFormulasBeforeSaving, recalculateFormulasOnOpening,
                keepTemplateSheet, areaBuilder, commands, clearTemplateCells, transformerFactory, streaming, template);
    }
    
    /**
     * @param data -
     * @param output -
     * @throws IOException 
     */
    public void buildAndFill(Map<String, Object> data, JxlsOutput output) {
        build().fill(data, output);
    }

    /**
     * @param data -
     * @param outputFile -
     */
	public void buildAndFill(Map<String, Object> data, File outputFile) {
		buildAndFill(data, new JxlsOutputFile(outputFile));
	}

    public SELF withExpressionEvaluatorFactory(ExpressionEvaluatorFactory expressionEvaluatorFactory) {
        if (expressionEvaluatorFactory == null) {
            throw new IllegalArgumentException("expressionEvaluatorFactory must not be null");
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

    public SELF withExceptionHandler(ExceptionHandler exceptionHandler) {
    	this.exceptionHandler = exceptionHandler;
    	return (SELF) this;
    }

	public SELF withFormulaProcessor(FormulaProcessor formulaProcessor) {
	    this.formulaProcessor = formulaProcessor; // can be null
	    return (SELF) this;
	}

	public FormulaProcessor getFormulaProcessor() {
	    return formulaProcessor;
	}

	public SELF withFastFormulaProcessor() {
	    return withFormulaProcessor(new FastFormulaProcessor());
	}
	
	public SELF withIgnoreColumnProps(boolean ignoreColumnProps) {
		this.ignoreColumnProps = ignoreColumnProps;
		return (SELF) this;
	}
	
	public SELF withIgnoreRowProps(boolean ignoreRowProps) {
		this.ignoreRowProps = ignoreRowProps;
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

    public SELF withKeepTemplateSheet(KeepTemplateSheet keepTemplateSheet) {
        this.keepTemplateSheet = keepTemplateSheet;
        return (SELF) this;
    }

    public SELF withHideTemplateSheet() {
        keepTemplateSheet = KeepTemplateSheet.HIDE;
        return (SELF) this;
    }
    
    public SELF withAreaBuilder(AreaBuilder areaBuilder) {
    	if (areaBuilder == null) {
    		throw new IllegalArgumentException("areaBuilder must not be null");
    	}
        this.areaBuilder = areaBuilder;
        return (SELF) this;
    }

    public AreaBuilder getAreaBuilder() {
        return areaBuilder;
    }
    
    public SELF withCommand(String name, Class<? extends Command> commandClass) {
    	commands.put(name, commandClass);
    	return (SELF) this;
    }

    public SELF withClearTemplateCells(boolean clearTemplateCells) {
        this.clearTemplateCells = clearTemplateCells;
        return (SELF) this;
    }
    
    public SELF withTransformerFactory(JxlsTransformerFactory transformerFactory) {
    	if (transformerFactory == null) {
    		throw new IllegalArgumentException("transformerFactory must not be null");
    	}
        this.transformerFactory = transformerFactory;
        return (SELF) this;
    }

    public JxlsTransformerFactory getTransformerFactory() {
        return transformerFactory;
    }

    public SELF withStreaming(JxlsStreaming streaming) {
        this.streaming = streaming == null ? JxlsStreaming.STREAMING_OFF : streaming;
        return (SELF) this;
    }

    public SELF withTemplate(InputStream template) {
    	if (template == null) {
    		throw new CannotOpenWorkbookException();
    	}
        this.template = template;
        return (SELF) this;
    }

    public SELF withTemplate(URL template) throws IOException {
    	if (template == null) {
    		throw new IllegalArgumentException("template must not be null");
    	}
        return withTemplate(template.openStream());
    }

    public SELF withTemplate(File template) throws FileNotFoundException {
    	if (template == null) {
    		throw new IllegalArgumentException("template must not be null");
    	} else if (!template.isFile()) {
    		throw new JxlsException("Template file does not exist: " + template.getAbsolutePath());
    	}
        return withTemplate(new FileInputStream(template));
    }

    public SELF withTemplate(String templateFileName) throws FileNotFoundException {
    	if (templateFileName == null || templateFileName.isBlank()) {
    		throw new IllegalArgumentException("Please specify templateFileName");
    	}
        return withTemplate(new File(templateFileName));
    }
}
