package org.jxls.builder;

import java.io.ByteArrayOutputStream;
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
import org.jxls.common.RunVarAccess;
import org.jxls.expression.ExpressionEvaluatorFactory;
import org.jxls.expression.ExpressionEvaluatorFactoryJexlImpl;
import org.jxls.formula.FastFormulaProcessor;
import org.jxls.formula.FormulaProcessor;
import org.jxls.formula.StandardFormulaProcessor;
import org.jxls.logging.JxlsLogger;
import org.jxls.transform.ExpressionEvaluatorContext;
import org.jxls.transform.ExpressionEvaluatorContextFactory;
import org.jxls.transform.JxlsTransformerFactory;
import org.jxls.transform.PreWriteAction;
import org.jxls.util.CannotOpenWorkbookException;

/**
 * This builder is the starting point for creating reports using Jxls.
 * After setting the options using the with... methods, call build(). This results in a TemplateFiller with a fill() method.
 * See website for detailed explanations of all options.
 * You must call withTransformerFactory() and withTemplate().
 */
public class JxlsTemplateFillerBuilder<SELF extends JxlsTemplateFillerBuilder<SELF>> {
    protected ExpressionEvaluatorContextFactory expressionEvaluatorContextFactory = (f, b, e) -> new ExpressionEvaluatorContext(f, b, e);
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
    protected RunVarAccess runVarAccess;

    /**
     * @return new builder instance
     */
    public static JxlsTemplateFillerBuilder<?> newInstance() {
        return new JxlsTemplateFillerBuilder<>();
    }

    /**
     * @return JxlsTemplateFiller with all options and the template
     */
    public JxlsTemplateFiller build() {
        if (logger == null) {
            throw new JxlsException("Please call withLogger()");
        } else if (transformerFactory == null) {
    		throw new JxlsException("Please call withTransformerFactory()");
    	} else if (template == null) {
    		throw new JxlsException("Please call withTemplate()");
    	}
        return new JxlsTemplateFiller(getOptions(), template);
    }
    
    /**
     * Builds template filler, creates Excel report with given data and saves it to output.
     * @param data not null
     * @param output not null
     */
    public void buildAndFill(Map<String, Object> data, JxlsOutput output) {
        build().fill(data, output);
    }

    /**
     * Builds template filler, creates Excel report with given data and saves it to given output file.
     * @param data -
     * @param outputFile -
     */
	public void buildAndFill(Map<String, Object> data, File outputFile) {
		buildAndFill(data, new JxlsOutputFile(outputFile));
	}

	/**
     * Builds template filler, creates Excel report with given data and returns Excel file as byte array.
	 * @param data -
	 * @return byte[]
	 */
    public byte[] buildAndFill(Map<String, Object> data) {
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        buildAndFill(data, () -> os);
        return os.toByteArray();
    }

    /**
     * Creates internal used options object for template filler
     * @return JxlsOptions
     */
    public JxlsOptions getOptions() {
        return new JxlsOptions(expressionEvaluatorContextFactory, expressionEvaluatorFactory, expressionNotationBegin, expressionNotationEnd,
                logger, formulaProcessor, updateCellDataArea, ignoreColumnProps, ignoreRowProps,
                recalculateFormulasBeforeSaving, recalculateFormulasOnOpening, keepTemplateSheet,
                areaBuilder, commands, clearTemplateCells, transformerFactory, streaming, needsContextList,
                preWriteActions, runVarAccess);
    }
    
    /**
     * Use this method for defining your own factory for building the ExpressionEvaluatorContext instance.
     * This is typically used for exchanging the ExpressionEvaluatorContext.evaluateRawExpression() method.
     * 
     * @param expressionEvaluatorContextFactory not null
     * @return this
     */
    public SELF withExpressionEvaluatorContextFactory(ExpressionEvaluatorContextFactory expressionEvaluatorContextFactory) {
        if (expressionEvaluatorContextFactory == null) {
            throw new IllegalArgumentException("expressionEvaluatorContextFactory must not be null");
        }
        this.expressionEvaluatorContextFactory = expressionEvaluatorContextFactory;
        return (SELF) this;
    }

    /**
     * Defines a factory class with which ExpressionEvaluator classes are created during report creation. It is recommended to use
     * <code>new ExpressionEvaluatorFactoryJexlImpl(false, true, JxlsJexlPermissions.RESTRICTED)</code> in production use.
     * 
     * @param expressionEvaluatorFactory not null. Default is ExpressionEvaluatorFactoryJexlImpl with settings: silent=true, strict=false,
     *                                   JxlsJexlPermissions.UNRESTRICTED.
     * @return this
     */
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

    /**
     * Expressions in Excel cells are inside <code>${</code> and <code>}</code>. Use this method to change those Strings.
     * @param begin You can use null for the default value "${".
     * @param end You can use null for the default value "}".
     * @return this
     */
    public SELF withExpressionNotation(String begin, String end) {
	    expressionNotationBegin = begin == null ? DEFAULT_EXPRESSION_BEGIN : begin;
	    expressionNotationEnd = end == null ? DEFAULT_EXPRESSION_END : end;
	    return (SELF) this;
	}

    /**
     * Defines which class should be used for logging and throwing exceptions.
     * @param logger not null
     * @return this
     */
    public SELF withLogger(JxlsLogger logger) {
        if (logger == null) {
            throw new IllegalArgumentException("logger must not be null");
        }
    	this.logger = logger;
    	return (SELF) this;
    }

    /**
     * @param formulaProcessor null for disabling formula processing, default: StandardFormulaProcessor
     * @return this
     */
	public SELF withFormulaProcessor(FormulaProcessor formulaProcessor) {
	    this.formulaProcessor = formulaProcessor; // can be null
	    return (SELF) this;
	}

	public FormulaProcessor getFormulaProcessor() {
	    return formulaProcessor;
	}

	/**
	 * Sets FastFormulaProcessor which is 10 times faster but can handle only simple templates.
	 * @return this
	 */
	public SELF withFastFormulaProcessor() {
	    return withFormulaProcessor(new FastFormulaProcessor());
	}
	
	/**
	 * Cell reference tracking
	 * @param updateCellDataArea false to turn this feature off
	 * @return this
	 */
	public SELF withUpdateCellDataArea(boolean updateCellDataArea) {
	    this.updateCellDataArea = updateCellDataArea;
	    return (SELF) this;
	}
	
	/**
     * See website
	 * @param ignoreColumnProps default is false
	 * @return this
	 */
	public SELF withIgnoreColumnProps(boolean ignoreColumnProps) {
		this.ignoreColumnProps = ignoreColumnProps;
		return (SELF) this;
	}
	
	/**
	 * See website
	 * @param ignoreRowProps default is false
	 * @return this
	 */
	public SELF withIgnoreRowProps(boolean ignoreRowProps) {
		this.ignoreRowProps = ignoreRowProps;
		return (SELF) this;
	}

	/**
	 * Call this method with argument false if you don't need an extra recalculation of all formula results before saving the Excel file.
	 * @param recalculateFormulasBeforeSaving default is true
	 * @return this
	 */
	public SELF withRecalculateFormulasBeforeSaving(boolean recalculateFormulasBeforeSaving) {
        this.recalculateFormulasBeforeSaving = recalculateFormulasBeforeSaving;
        return (SELF) this;
    }

	/**
	 * Call this method with argument true if you all formula results should be recalculated
	 * while opening the report file using Microsoft Excel.
	 * @param recalculateFormulasOnOpening default is false
	 * @return this
	 */
    public SELF withRecalculateFormulasOnOpening(boolean recalculateFormulasOnOpening) {
        this.recalculateFormulasOnOpening = recalculateFormulasOnOpening;
        return (SELF) this;
    }

    /**
     * This is for the multisheet feature. The template sheet is deleted by default.
     * Call this method with argument KEEP if you want to keep the template sheet.
     * Call this method with argument HIDE if you just want to hide the template sheet.
     * @param keepTemplateSheet default is DELETE, null will be ignored.
     * @return this
     */
    public SELF withKeepTemplateSheet(KeepTemplateSheet keepTemplateSheet) {
        if (keepTemplateSheet != null) {
            this.keepTemplateSheet = keepTemplateSheet;
        }
        return (SELF) this;
    }

    /**
     * See website
     * @param areaBuilder not null. Default is XlsCommentAreaBuilder.
     * @return this
     */
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
    
    /**
     * Adds a command.
     * @param name command name without "jx": prefix
     * @param commandClass -
     * @return this
     */
    public SELF withCommand(String name, Class<? extends Command> commandClass) {
    	commands.put(name, commandClass);
    	return (SELF) this;
    }

    /**
     * @param clearTemplateCells true (default): cells where the expression can not be evaluated will be cleared,
     *                           false: cells where the expression can not be evaluated will keep the expression.
     * @return this
     */
    public SELF withClearTemplateCells(boolean clearTemplateCells) {
        this.clearTemplateCells = clearTemplateCells;
        return (SELF) this;
    }
    
    /**
     * See website
     * @param transformerFactory -
     * @return this
     */
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

    /**
     * Use streaming if you have large sheets. However, not all features can be used.
     * See website.
     * @param streaming Default is STREAMING_OFF for no streaming.
     *                  STREAMING_ON: streaming for all sheets.
     *                  AUTO_DETECT: streaming for sheets which have the exact text <code>sheetStreaming="true"</code> in its jx:area command.
     * @return this
     */
    public SELF withStreaming(JxlsStreaming streaming) {
        this.streaming = streaming == null ? JxlsStreaming.STREAMING_OFF : streaming;
        return (SELF) this;
    }
    
    public JxlsStreaming getStreaming() {
        return streaming;
    }
    
    /**
     * If you are adding an object to the data map which needs <code>PublicContext</code> access, implement the NeedsPublicContext interface
     * and call this method.
     * @param needsPublicContext your class which implements NeedsPublicContext for getting the PublicContext instance during report creation
     * @return this
     */
    public SELF needsPublicContext(NeedsPublicContext needsPublicContext) {
        if (needsPublicContext == null) {
            throw new IllegalArgumentException("needsPublicContext must not be null");
        }
        needsContextList.add(needsPublicContext);
        return (SELF) this;
    }
    
    /**
     * @param preWriteAction code to be executed before transformer.write() is executed 
     * @return this
     */
    public SELF withPreWriteAction(PreWriteAction preWriteAction) {
        if (preWriteAction == null) {
            throw new IllegalArgumentException("preWriteAction must not be null");
        }
        preWriteActions.add(preWriteAction);
        return (SELF) this;
    }

    /**
     * Change the behavior for accessing run vars.
     * @param runVarAccess code that retrieves the value for a given key from the given data map
     * @return this
     */
    public SELF withRunVarAccess(RunVarAccess runVarAccess) {
        this.runVarAccess = runVarAccess;
        return (SELF) this;
    }

    /**
     * @param templateInputStream Excel template as InputStream
     * @return this
     */
    public SELF withTemplate(InputStream templateInputStream) {
    	if (templateInputStream == null) {
    		throw new CannotOpenWorkbookException();
    	}
        template = templateInputStream;
        return (SELF) this;
    }

    /**
     * @param templateURL Excel template as file defined by an URL
     * @return this
     * @throws IOException
     */
    public SELF withTemplate(URL templateURL) throws IOException {
    	if (templateURL == null) {
    		throw new IllegalArgumentException("templateURL must not be null");
    	}
        return withTemplate(templateURL.openStream());
    }

    /**
     * @param templateFile Excel template file
     * @return this
     * @throws FileNotFoundException
     */
    public SELF withTemplate(File templateFile) throws FileNotFoundException {
    	if (templateFile == null) {
    		throw new IllegalArgumentException("templateFile must not be null");
    	} else if (!templateFile.isFile()) {
    		throw new JxlsException("Template file does not exist: " + templateFile.getAbsolutePath());
    	}
        return withTemplate(new FileInputStream(templateFile));
    }

    /**
     * @param templateFileName Excel template file name
     * @return this
     * @throws FileNotFoundException
     */
    public SELF withTemplate(String templateFileName) throws FileNotFoundException {
    	if (templateFileName == null || templateFileName.isBlank()) {
    		throw new IllegalArgumentException("Please specify templateFileName");
    	}
        return withTemplate(new File(templateFileName));
    }
}
