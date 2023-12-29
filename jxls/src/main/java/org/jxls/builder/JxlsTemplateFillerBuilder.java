package org.jxls.builder;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.command.Command;
import org.jxls.common.JxlsException;
import org.jxls.common.NeedsPublicContext;
import org.jxls.expression.ExpressionEvaluatorFactory;
import org.jxls.expression.ExpressionEvaluatorFactoryJexlImpl;
import org.jxls.formula.FastFormulaProcessor;
import org.jxls.formula.FormulaProcessor;
import org.jxls.formula.StandardFormulaProcessor;
import org.jxls.logging.JxlsLogger;
import org.jxls.transform.JxlsTransformerFactory;
import org.jxls.transform.PreWriteAction;
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
    protected JxlsLogger logger;
    /** null: no formula processing */
    private FormulaProcessor formulaProcessor = new StandardFormulaProcessor();
    /** old name: formulaProcessingRequired */
    protected boolean updateCellDataArea = true;
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
    private JxlsStreaming streaming = JxlsStreaming.STREAMING_OFF;
    protected InputStream template;
    protected final List<NeedsPublicContext> needsContextList = new ArrayList<>();
    protected final List<PreWriteAction> preWriteActions = new ArrayList<>();

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
        return new JxlsTemplateFiller(getOptions(), template);
    }
    
    /**
     * @param data -
     * @param output -
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

    public JxlsOptions getOptions() {
        return new JxlsOptions(expressionEvaluatorFactory, expressionNotationBegin, expressionNotationEnd,
                logger, formulaProcessor, updateCellDataArea, ignoreColumnProps, ignoreRowProps, recalculateFormulasBeforeSaving, recalculateFormulasOnOpening,
                keepTemplateSheet, areaBuilder, commands, clearTemplateCells, transformerFactory, streaming, needsContextList, preWriteActions);
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

    public SELF withLogger(JxlsLogger logger) {
        if (logger == null) {
            throw new IllegalArgumentException("logger must not be null");
        }
    	this.logger = logger;
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
	
	public SELF withUpdateCellDataArea(boolean updateCellDataArea) {
	    this.updateCellDataArea = updateCellDataArea;
	    return (SELF) this;
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
    
    public JxlsStreaming getStreaming() {
        return streaming;
    }
    
    public SELF needsPublicContext(NeedsPublicContext needsPublicContext) {
        if (needsPublicContext == null) {
            throw new IllegalArgumentException("needsPublicContext must not be null");
        }
        needsContextList.add(needsPublicContext);
        return (SELF) this;
    }
    
    public SELF withPreWriteAction(PreWriteAction preWriteAction) {
        if (preWriteAction == null) {
            throw new IllegalArgumentException("preWriteAction must not be null");
        }
        preWriteActions.add(preWriteAction);
        return (SELF) this;
    }

    public SELF withTemplate(InputStream templateInputStream) {
    	if (templateInputStream == null) {
    		throw new CannotOpenWorkbookException();
    	}
        template = templateInputStream;
        return (SELF) this;
    }

    public SELF withTemplate(URL templateURL) throws IOException {
    	if (templateURL == null) {
    		throw new IllegalArgumentException("templateURL must not be null");
    	}
        return withTemplate(templateURL.openStream());
    }

    public SELF withTemplate(File templateFile) throws FileNotFoundException {
    	if (templateFile == null) {
    		throw new IllegalArgumentException("templateFile must not be null");
    	} else if (!templateFile.isFile()) {
    		throw new JxlsException("Template file does not exist: " + templateFile.getAbsolutePath());
    	}
        return withTemplate(new FileInputStream(templateFile));
    }

    public SELF withTemplate(String templateFileName) throws FileNotFoundException {
    	if (templateFileName == null || templateFileName.isBlank()) {
    		throw new IllegalArgumentException("Please specify templateFileName");
    	}
        return withTemplate(new File(templateFileName));
    }
}
