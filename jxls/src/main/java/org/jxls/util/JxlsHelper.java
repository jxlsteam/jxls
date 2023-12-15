package org.jxls.util;

import static org.jxls.util.Util.getSheetsNameOfMultiSheetTemplate;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Collection;
import java.util.List;

import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.command.GridCommand;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.expression.ExpressionEvaluatorFactory;
import org.jxls.expression.ExpressionEvaluatorFactoryJexlImpl;
import org.jxls.formula.FastFormulaProcessor;
import org.jxls.formula.FormulaProcessor;
import org.jxls.formula.StandardFormulaProcessor;
import org.jxls.template.SimpleExporter;
import org.jxls.transform.Transformer;

/**
 * Helper class to simplify JXLS usage
 */
public class JxlsHelper {
    private static final ExpressionEvaluatorFactory INSTANCE = new ExpressionEvaluatorFactoryJexlImpl();
    private boolean hideTemplateSheet = false;
    private boolean deleteTemplateSheet = true;
    private boolean processFormulas = true;
    private boolean useFastFormulaProcessor = false;
    private boolean evaluateFormulas = false;
    private boolean fullFormulaRecalculationOnOpening = false;
    private String expressionNotationBegin;
    private String expressionNotationEnd;
    private FormulaProcessor formulaProcessor;
    private SimpleExporter simpleExporter = new SimpleExporter();
    private AreaBuilder areaBuilder = new XlsCommentAreaBuilder();
    private boolean clearTemplateCells = true;

    public JxlsHelper() {
    }

    /**
     * Returns new JxlsHelper instance
     * @return JxlsHelper object
     */
    public static JxlsHelper getInstance() {
        return new JxlsHelper();
    }

    /**
     * This property is used to recalculate all formulas before saving the workbook.
     * This property is set to true if you don't open the file with MS Excel and just read it (e.g. with a unit test).
     * The following documentation is POI specific.
     * 
     * @return true: calls <code>workbook.getCreationHelper().createFormulaEvaluator().evaluateAll()</code>
     * before writing the workbook. Please have a look at the POI documentation for more details.
     * This does not work for streaming.
     * Please be aware that POI supports only a subset of Excel formulas.
     * If an unsupported formula is in the template the evaluation will fail.
     * <p>false (default): do nothing (hopefully MS Excel will recalculate all formulas while opening the file)</p>
     */
    public boolean isEvaluateFormulas() {
        return evaluateFormulas;
    }

    /**
     * This property is used to recalculate all formulas before saving the workbook.
     * This property is set to true if you don't open the file with MS Excel and just read it (e.g. with a unit test).
     * 
     * <p>Don't do this:<br>
     * <code>JxlsHelper.getInstance().setEvaluateFormulas(true);<br>
     * JxlsHelper.getInstance().processTemplate(...);</code>
     * <br>Call JxlsHelper.getInstance() only once!</p>
     * 
     * <p>The following documentation is POI specific.</p>
     * 
     * @param evaluateFormulas
     * true: calls <code>workbook.getCreationHelper().createFormulaEvaluator().evaluateAll()</code>
     * before writing the workbook. Please have a look at the POI documentation for more details.
     * This does not work for streaming.
     * Please be aware that POI supports only a subset of Excel formulas.
     * If an unsupported formula is in the template the evaluation will fail.
     * <p>false (default): do nothing (hopefully MS Excel will recalculate all formulas while opening the file)</p>
     * @return this
     */
    public JxlsHelper setEvaluateFormulas(boolean evaluateFormulas) {
        this.evaluateFormulas = evaluateFormulas;
        return this;
    }

    /**
     * @return false: do nothing, true: activate recalculation when opening
     */
    public boolean isFullFormulaRecalculationOnOpening() {
        return fullFormulaRecalculationOnOpening;
    }

    /**
     * If you set this option to true, all formulas will be recalculated when the file is opened in MS Excel.
     * This changes the Excel file. This only works once. 
     * @param fullFormulaRecalculationOnOpening false: do nothing (default value),
     *        true: activate recalculation when opening
     * @return this
     */
    public JxlsHelper setFullFormulaRecalculationOnOpening(boolean fullFormulaRecalculationOnOpening) {
        this.fullFormulaRecalculationOnOpening = fullFormulaRecalculationOnOpening;
        return this;
    }

    public AreaBuilder getAreaBuilder() {
        return areaBuilder;
    }

    public JxlsHelper setAreaBuilder(AreaBuilder areaBuilder) {
        this.areaBuilder = areaBuilder;
        return this;
    }

    public void setClearTemplateCells(boolean clearTemplateCells) {
        this.clearTemplateCells = clearTemplateCells;
    }

    /**
     * @return current {@link FormulaProcessor} implementation
     */
    public FormulaProcessor getFormulaProcessor() {
        return formulaProcessor;
    }

    /**
     * Sets formula processor implementation
     *
     * @param formulaProcessor FormulaProcessor instance
     * @return this
     */
    public JxlsHelper setFormulaProcessor(FormulaProcessor formulaProcessor) {
        this.formulaProcessor = formulaProcessor;
        return this;
    }

    /**
     * @return true if formula processing is on
     */
    public boolean isProcessFormulas() {
        return processFormulas;
    }

    /**
     * Enables or disables formula processing
     *
     * @param processFormulas enable/disable formula processing flag
     * @return this
     */
    public JxlsHelper setProcessFormulas(boolean processFormulas) {
        this.processFormulas = processFormulas;
        return this;
    }

    /**
     * @return true if template sheet must be hidden
     */
    public boolean isHideTemplateSheet() {
        return hideTemplateSheet;
    }

    /**
     * Hides/shows template sheet
     *
     * @param hideTemplateSheet true to hide template sheet or false to show it
     * @return this
     */
    public JxlsHelper setHideTemplateSheet(boolean hideTemplateSheet) {
        this.hideTemplateSheet = hideTemplateSheet;
        return this;
    }

    /**
     * @return flag indicating if template sheet must be deleted
     */
    public boolean isDeleteTemplateSheet() {
        return deleteTemplateSheet;
    }

    /**
     * Marks template sheet for deletion
     *
     * @param deleteTemplateSheet true if template sheet should be deleted
     * @return this
     */
    public JxlsHelper setDeleteTemplateSheet(boolean deleteTemplateSheet) {
        this.deleteTemplateSheet = deleteTemplateSheet;
        return this;
    }

    /**
     * @return true if {@link FastFormulaProcessor} should be used
     */
    public boolean isUseFastFormulaProcessor() {
        return useFastFormulaProcessor;
    }

    /**
     * @param useFastFormulaProcessor
     *            if true the {@link FastFormulaProcessor} will be used to process formulas
     *            otherwise {@link StandardFormulaProcessor} will be used
     * @return this
     */
    public JxlsHelper setUseFastFormulaProcessor(boolean useFastFormulaProcessor) {
        this.useFastFormulaProcessor = useFastFormulaProcessor;
        return this;
    }

    /**
     * Allows to set custom notation for expressions in the template The notation is
     * used in {@link JxlsHelper#createTransformer(InputStream, OutputStream)}
     *
     * @param expressionNotationBegin notation prefix
     * @param expressionNotationEnd notation suffix
     * @return this
     */
    public JxlsHelper buildExpressionNotation(String expressionNotationBegin, String expressionNotationEnd) {
        this.expressionNotationBegin = expressionNotationBegin;
        this.expressionNotationEnd = expressionNotationEnd;
        return this;
    }

    /**
     * Reads template from the source stream processes it using the supplied
     * {@link Context} and writes the result to the target stream
     *
     * @param templateStream source input stream with the template
     * @param targetStream target output stream
     * @param context data map
     * @return this
     * @throws IOException -
     */
    public JxlsHelper processTemplate(InputStream templateStream, OutputStream targetStream, Context context)
            throws IOException {
        Transformer transformer = createTransformer(templateStream, targetStream);
        processTemplate(context, transformer);
        return this;
    }

    /**
     * Processes a template with the given {@link Transformer} instance and {@link Context}
     *
     * @param context data context
     * @param transformer transformer instance
     * @throws IOException -
     */
    public void processTemplate(Context context, Transformer transformer) throws IOException {
        List<Area> xlsAreaList = areaBuilder.build(transformer, clearTemplateCells);
        for (Area xlsArea : xlsAreaList) {
            xlsArea.applyAt(new CellRef(xlsArea.getStartCellRef().getCellName()), context);
        }
        if (processFormulas) {
            for (Area xlsArea : xlsAreaList) {
                setFormulaProcessor(xlsArea);
                xlsArea.processFormulas();
            }
        }
        if (isHideTemplateSheet()) {
            List<String> sheetNameTemplate = getSheetsNameOfMultiSheetTemplate(xlsAreaList);
            for (String sheetName : sheetNameTemplate) {
                transformer.setHidden(sheetName, true);
            }
        }
        if (isDeleteTemplateSheet()) {
            List<String> sheetNameTemplate = getSheetsNameOfMultiSheetTemplate(xlsAreaList);
            for (String sheetName : sheetNameTemplate) {
                transformer.deleteSheet(sheetName);
            }
        }
        transformer.write();
    }

    /**
     * @return current {@link ExpressionEvaluatorFactory} implementation
     */
    public ExpressionEvaluatorFactory getExpressionEvaluatorFactory() {
        return INSTANCE;
    }

    private Area setFormulaProcessor(Area xlsArea) {
        FormulaProcessor fp = formulaProcessor;
        if (fp == null) {
            if (useFastFormulaProcessor) {
                fp = new FastFormulaProcessor();
            } else {
                fp = new StandardFormulaProcessor();
            }
        }
        xlsArea.setFormulaProcessor(fp);
        return xlsArea;
    }

    /**
     * Processes the template from the given input stream using the supplied
     * {@link Context} and the given target cell and writes the result to the output stream
     *
     * @param templateStream source input stream for the template
     * @param targetStream target output stream to write the processing result
     * @param context data map
     * @param targetCell starting target cell into which the template area must be processed
     * @return this
     * @throws IOException if input/output stream processing resolves in an error
     */
    public JxlsHelper processTemplateAtCell(InputStream templateStream, OutputStream targetStream, Context context,
            String targetCell) throws IOException {
        Transformer transformer = createTransformer(templateStream, targetStream);
        List<Area> xlsAreaList = areaBuilder.build(transformer, clearTemplateCells);
        if (xlsAreaList.isEmpty()) {
            throw new IllegalStateException("No XlsArea were detected for this processing");
        }
        Area firstArea = xlsAreaList.get(0);
        CellRef targetCellRef = new CellRef(targetCell);
        firstArea.applyAt(targetCellRef, context);
        if (processFormulas) {
            setFormulaProcessor(firstArea);
            firstArea.processFormulas();
        }
        String sourceSheetName = firstArea.getStartCellRef().getSheetName();
        if (!sourceSheetName.equalsIgnoreCase(targetCellRef.getSheetName())) {
            if (hideTemplateSheet) {
                transformer.setHidden(sourceSheetName, true);
            }
            if (deleteTemplateSheet) {
                transformer.deleteSheet(sourceSheetName);
            }
        }
        transformer.write();
        return this;
    }

    /**
     * Processes the template with the {@link GridCommand}
     *
     * @param templateStream template input stream
     * @param targetStream output stream for the result
     * @param context data map
     * @param objectProps object properties to use with the {@link GridCommand}
     * @return this
     * @throws IOException -
     */
    public JxlsHelper processGridTemplate(InputStream templateStream, OutputStream targetStream, Context context, String objectProps) throws IOException {
        Transformer transformer = createTransformer(templateStream, targetStream);
        List<Area> xlsAreaList = areaBuilder.build(transformer, clearTemplateCells);
        for (Area xlsArea : xlsAreaList) {
            GridCommand gridCommand = (GridCommand) xlsArea.getCommandDataList().get(0).getCommand();
            gridCommand.setProps(objectProps);
            setFormulaProcessor(xlsArea);
            xlsArea.applyAt(new CellRef(xlsArea.getStartCellRef().getCellName()), context);
            if (processFormulas) {
                xlsArea.processFormulas();
            }
        }
        transformer.write();
        return this;
    }

    /**
     * Processes the input template with {@link GridCommand} at given target cell
     * using the given object properties and context
     *
     * @param templateStream template input stream
     * @param targetStream result output stream
     * @param context data map
     * @param objectProps object properties to use with {@link GridCommand}
     * @param targetCell start target cell to use when processing the template
     * @throws IOException -
     */
    public void processGridTemplateAtCell(InputStream templateStream, OutputStream targetStream, Context context,
            String objectProps, String targetCell) throws IOException {
        Transformer transformer = createTransformer(templateStream, targetStream);
        List<Area> xlsAreaList = areaBuilder.build(transformer, clearTemplateCells);
        Area firstArea = xlsAreaList.get(0);
        CellRef targetCellRef = new CellRef(targetCell);
        GridCommand gridCommand = (GridCommand) firstArea.getCommandDataList().get(0).getCommand();
        gridCommand.setProps(objectProps);
        firstArea.applyAt(targetCellRef, context);
        if (processFormulas) {
            setFormulaProcessor(firstArea);
            firstArea.processFormulas();
        }
        String sourceSheetName = firstArea.getStartCellRef().getSheetName();
        if (!sourceSheetName.equalsIgnoreCase(targetCellRef.getSheetName())) {
            if (hideTemplateSheet) {
                transformer.setHidden(sourceSheetName, true);
            }
            if (deleteTemplateSheet) {
                transformer.deleteSheet(sourceSheetName);
            }
        }
        transformer.write();
    }

    /**
     * Registers grid template to be used with {@link SimpleExporter}
     *
     * @param inputStream template input stream
     * @return this
     * @throws IOException -
     */
    public JxlsHelper registerGridTemplate(InputStream inputStream) throws IOException {
        simpleExporter.registerGridTemplate(inputStream);
        return this;
    }

    /**
     * Performs data grid export using {@link SimpleExporter}
     *
     * @param headers collection of headers to use for the export
     * @param dataObjects collection of data objects to export
     * @param objectProps object properties (comma separated) to use with the {@link GridCommand}
     * @param outputStream output stream to write the processing result
     * @return this
     */
    public JxlsHelper gridExport(Collection<?> headers, Collection<?> dataObjects, String objectProps, OutputStream outputStream) {
        simpleExporter.gridExport(headers, dataObjects, objectProps, outputStream);
        return this;
    }

    /**
     * Creates {@link Transformer} instance connected to the given input stream and
     * output stream with the default or custom expression notation (see also
     * {@link JxlsHelper#buildExpressionNotation(String, String)}
     *
     * @param templateStream source template input stream
     * @param targetStream target output stream to write the processing result
     * @return created transformer
     */
    public Transformer createTransformer(InputStream templateStream, OutputStream targetStream) {
        Transformer transformer = TransformerFactory.createTransformer(templateStream, targetStream);
        if (transformer == null) {
            throw new IllegalStateException("Cannot load XLS transformer. Please make sure a Transformer implementation is in classpath");
        }
        if (expressionNotationBegin != null && expressionNotationEnd != null) {
            transformer.getTransformationConfig().buildExpressionNotation(expressionNotationBegin,
                    expressionNotationEnd);
        }
        transformer.setEvaluateFormulas(evaluateFormulas);
        transformer.setFullFormulaRecalculationOnOpening(fullFormulaRecalculationOnOpening);
        return transformer;
    }
}
