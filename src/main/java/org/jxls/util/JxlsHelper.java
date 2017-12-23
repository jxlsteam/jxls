package org.jxls.util;

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
import org.jxls.expression.ExpressionEvaluator;
import org.jxls.expression.ExpressionEvaluatorFactory;
import org.jxls.formula.FastFormulaProcessor;
import org.jxls.formula.FormulaProcessor;
import org.jxls.formula.StandardFormulaProcessor;
import org.jxls.template.SimpleExporter;
import org.jxls.transform.Transformer;

/** Helper class to simplify Jxls usage */
public class JxlsHelper {
  private boolean hideTemplateSheet = false;
  private boolean deleteTemplateSheet = true;
  private boolean processFormulas = true;
  private boolean useFastFormulaProcessor = true;
  private String expressionNotationBegin;
  private String expressionNotationEnd;
  private FormulaProcessor formulaProcessor;
  private SimpleExporter simpleExporter = new SimpleExporter();

  private AreaBuilder areaBuilder = new XlsCommentAreaBuilder();

  private static final ServiceFactory SERVICE_FACTORY =
      ServiceFactory.DEFAULT.createService(ServiceFactory.class, ServiceFactory.DEFAULT);

  private static <T> T loadService(Class<T> interfaceClass) {
    final T ret = SERVICE_FACTORY.createService(interfaceClass, null);
    return ret;
  }

  private static final class ExpressionEvaluatorFactoryHolder {
    private static final ExpressionEvaluatorFactory INSTANCE;

    static {
      INSTANCE = loadService(ExpressionEvaluatorFactory.class);
    }
  }

  private static final JxlsConfigProvider CONFIG_PROVIDER = loadService(JxlsConfigProvider.class);

  public static JxlsHelper getInstance() {
    return new JxlsHelper();
  }

  public JxlsHelper() {}

  public AreaBuilder getAreaBuilder() {
    return areaBuilder;
  }

  public JxlsHelper setAreaBuilder(AreaBuilder areaBuilder) {
    this.areaBuilder = areaBuilder;
    return this;
  }

  /** @return current {@link FormulaProcessor} implementation */
  public FormulaProcessor getFormulaProcessor() {
    return formulaProcessor;
  }

  /**
   * Sets formula processor implementation
   *
   * @param formulaProcessor
   * @return this JxlsHelper instance
   */
  public JxlsHelper setFormulaProcessor(FormulaProcessor formulaProcessor) {
    this.formulaProcessor = formulaProcessor;
    return this;
  }

  /** @return true if formula processing is on */
  public boolean isProcessFormulas() {
    return processFormulas;
  }

  /**
   * Enables or disables formula processing
   *
   * @param processFormulas enable/disable formula processing flag
   * @return this JxlsHelper instance
   */
  public JxlsHelper setProcessFormulas(boolean processFormulas) {
    this.processFormulas = processFormulas;
    return this;
  }

  /** @return true if template sheet must be hidden */
  public boolean isHideTemplateSheet() {
    return hideTemplateSheet;
  }

  /**
   * Hides/shows template sheet
   *
   * @param hideTemplateSheet true to hide template sheet or false to show it
   * @return this JxlsHelper instance
   */
  public JxlsHelper setHideTemplateSheet(boolean hideTemplateSheet) {
    this.hideTemplateSheet = hideTemplateSheet;
    return this;
  }

  /** @return flag indicating if template sheet must be deleted */
  public boolean isDeleteTemplateSheet() {
    return deleteTemplateSheet;
  }

  /**
   * Marks template sheet for deletion
   *
   * @param deleteTemplateSheet true if template sheet should be deleted
   * @return this JxlsHelper instance
   */
  public JxlsHelper setDeleteTemplateSheet(boolean deleteTemplateSheet) {
    this.deleteTemplateSheet = deleteTemplateSheet;
    return this;
  }

  /** @return true if {@link FastFormulaProcessor} should be used */
  public boolean isUseFastFormulaProcessor() {
    return useFastFormulaProcessor;
  }

  /**
   * @param useFastFormulaProcessor if true the {@link FastFormulaProcessor} will be used to process
   *     formulas otherwise {@link StandardFormulaProcessor} will be used
   * @return this JxlsHelper instance
   */
  public JxlsHelper setUseFastFormulaProcessor(boolean useFastFormulaProcessor) {
    this.useFastFormulaProcessor = useFastFormulaProcessor;
    return this;
  }

  /**
   * Allows to set custom notation for expressions in the template
   * The notation is used in {@link JxlsHelper#createTransformer(InputStream, OutputStream)}
   * @param expressionNotationBegin - notation prefix
   * @param expressionNotationEnd - notation suffix
   * @return this JxlsHelper instance
   */
  public JxlsHelper buildExpressionNotation(
      String expressionNotationBegin, String expressionNotationEnd) {
    this.expressionNotationBegin = expressionNotationBegin;
    this.expressionNotationEnd = expressionNotationEnd;
    return this;
  }

  /**
   * Reads template from the source stream processes it using the supplied {@link Context} and
   * writes the result to the target stream
   *
   * @param templateStream source input stream with the template
   * @param targetStream target output stream
   * @param context data map
   * @return this JxlsHelper instance
   * @throws IOException
   */
  public JxlsHelper processTemplate(
      InputStream templateStream, OutputStream targetStream, Context context) throws IOException {
    Transformer transformer = createTransformer(templateStream, targetStream);
    processTemplate(context, transformer);
    return this;
  }

  /**
   * Processes a template with the given {@link Transformer} instance and {@link Context}
   *
   * @param context data context
   * @param transformer transformer instance
   * @throws IOException
   */
  public void processTemplate(Context context, Transformer transformer) throws IOException {
    areaBuilder.setTransformer(transformer);
    List<Area> xlsAreaList = areaBuilder.build();
    for (Area xlsArea : xlsAreaList) {
      xlsArea.applyAt(new CellRef(xlsArea.getStartCellRef().getCellName()), context);
      if (processFormulas) {
        setFormulaProcessor(xlsArea);
        xlsArea.processFormulas();
      }
    }
    transformer.write();
  }

  /**
   * Returns the configuration property value
   *
   * @param key property key
   * @param defaultValue default value to use if undefined
   * @return property value or the passed default value if the property not found
   */
  public static String getProperty(final String key, final String defaultValue) {
    return CONFIG_PROVIDER.getProperty(key, defaultValue);
  }

  /** @return current {@link ExpressionEvaluatorFactory} implementation */
  public ExpressionEvaluatorFactory getExpressionEvaluatorFactory() {
    return ExpressionEvaluatorFactoryHolder.INSTANCE;
  }

  /**
   * Creates {@link ExpressionEvaluator} instance for evaluation of the given expression
   *
   * @param expression expression to evaluate
   * @return {@link ExpressionEvaluator} instance for evaluation the passed expression
   */
  public ExpressionEvaluator createExpressionEvaluator(final String expression) {
    return ExpressionEvaluatorFactoryHolder.INSTANCE.createExpressionEvaluator(expression);
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
     * Processes the template from the given input stream
     * using the supplied {@link Context} and the given target cell
     * and writes the result to the output stream
     * @param templateStream source input stream for the template
     * @param targetStream target output stream to write the processing result
     * @param context data context map
     * @param targetCell starting target cell into which the template area must be processed
     * @return this JxlsHelper instance
     * @throws IOException if input/output stream processing resolves in an error
     */
  public JxlsHelper processTemplateAtCell(
      InputStream templateStream, OutputStream targetStream, Context context, String targetCell)
      throws IOException {
    Transformer transformer = createTransformer(templateStream, targetStream);
    areaBuilder.setTransformer(transformer);
    List<Area> xlsAreaList = areaBuilder.build();
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
     * @param templateStream template input stream
     * @param targetStream output stream for the result
     * @param context context data map
     * @param objectProps object properties to use with the {@link GridCommand}
     * @return this JxlsHelper instance
     * @throws IOException
     */
  public JxlsHelper processGridTemplate(
      InputStream templateStream, OutputStream targetStream, Context context, String objectProps)
      throws IOException {
    Transformer transformer = createTransformer(templateStream, targetStream);
    areaBuilder.setTransformer(transformer);
    List<Area> xlsAreaList = areaBuilder.build();
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
     * @param templateStream template input stream
     * @param targetStream result output stream
     * @param context context data map
     * @param objectProps object properties to use with {@link GridCommand}
     * @param targetCell start target cell to use when processing the template
     * @throws IOException
     */
  public void processGridTemplateAtCell(
      InputStream templateStream,
      OutputStream targetStream,
      Context context,
      String objectProps,
      String targetCell)
      throws IOException {
    Transformer transformer = createTransformer(templateStream, targetStream);
    areaBuilder.setTransformer(transformer);
    List<Area> xlsAreaList = areaBuilder.build();
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
     * @param inputStream template input stream
     * @return this JxlsHelper object
     * @throws IOException
     */
  public JxlsHelper registerGridTemplate(InputStream inputStream) throws IOException {
    simpleExporter.registerGridTemplate(inputStream);
    return this;
  }

    /**
     * Performs data grid export using {@link SimpleExporter}
     * @param headers collection of headers to use for the export
     * @param dataObjects collection of data objects to export
     * @param objectProps object properties (comma separated) to use with the {@link GridCommand}
     * @param outputStream output stream to write the processing result
     * @return
     */
  public JxlsHelper gridExport(
      Collection headers, Collection dataObjects, String objectProps, OutputStream outputStream) {
    simpleExporter.gridExport(headers, dataObjects, objectProps, outputStream);
    return this;
  }

  /**
   * Creates {@link Transformer} instance connected to the given input stream
   * and output stream with the default or custom expression notation (see also {@link JxlsHelper#buildExpressionNotation(String, String)}
   * @param templateStream source template input stream
   * @param targetStream target output stream to write the processing result
   * @return
   */
  public Transformer createTransformer(InputStream templateStream, OutputStream targetStream) {
    Transformer transformer = TransformerFactory.createTransformer(templateStream, targetStream);
    if (transformer == null) {
      throw new IllegalStateException(
          "Cannot load XLS transformer. Please make sure a Transformer implementation is in classpath");
    }
    if (expressionNotationBegin != null && expressionNotationEnd != null) {
      transformer
          .getTransformationConfig()
          .buildExpressionNotation(expressionNotationBegin, expressionNotationEnd);
    }
    return transformer;
  }
}
