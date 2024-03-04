package org.jxls.builder;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;

import org.jxls.area.Area;
import org.jxls.command.Command;
import org.jxls.command.EachCommand;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.ContextImpl;
import org.jxls.common.RunVarAccess;
import org.jxls.transform.ExpressionEvaluatorContext;
import org.jxls.transform.Transformer;

public class JxlsTemplateFiller {
    protected final JxlsOptions options;
    protected final InputStream template;
    protected Transformer transformer;
    protected List<Area> areas;
    private Context context;
	
    protected JxlsTemplateFiller(JxlsOptions options, InputStream template) {
        this.options = options;
        this.template = template;
    }

    /**
     * Creates Excel report using the given data map and writes it to the given output.
     * @param data not null
     * @param output not null
     */
    public void fill(Map<String, Object> data, JxlsOutput output) {
        if (data == null || output == null) {
            throw new IllegalArgumentException();
        }
        try (OutputStream outputStream = output.getOutputStream()) {
            createTransformer(outputStream);
            configureTransformer();
            installCommands();
            processAreas(data);
            preWrite();
            write();
        } catch (IOException e) {
            throw new JxlsTemplateFillException(e);
        } finally {
            context = null;
            areas = null;
            transformer = null;
        }
    }

    protected void createTransformer(OutputStream outputStream) {
        transformer = options.getTransformerFactory().create(template, outputStream, options.getStreaming(), options.getLogger());
    }

    protected void configureTransformer() {
    	transformer.setIgnoreColumnProps(options.isIgnoreColumnProps());
    	transformer.setIgnoreRowProps(options.isIgnoreRowProps());
    	transformer.setSheetCreater(options.getSheetCreater());
    }

    private void installCommands() {
        if (options.getAreaBuilder() instanceof CommandMappings cm) {
            options.getCommands().forEach((k, v) -> cm.addCommandMapping(k, v));
        }
    }

    /**
     * Implementation must set areas variable.
     * @param data -
     */
    protected void processAreas(Map<String, Object> data) {
        areas = options.getAreaBuilder().build(transformer, options.isClearTemplateCells());
        
        context = createContext(createExpressionEvaluatorContext(), data, options.getRunVarAccess());
        context.setUpdateCellDataArea(options.isUpdateCellDataArea());
        options.getNeedsPublicContextList().forEach(ee -> ee.setPublicContext(context));
        
        areas.forEach(area -> area.applyAt(new CellRef(area.getStartCellRef().getCellName()), context));
        
        if (options.getFormulaProcessor() != null) {
            areas.forEach(area -> area.processFormulas(options.getFormulaProcessor()));
        }
    }

    protected Context createContext(ExpressionEvaluatorContext expressionEvaluatorContext, Map<String, Object> data, RunVarAccess runVarAccess) {
        return new ContextImpl(expressionEvaluatorContext, data, runVarAccess);
    }

    protected ExpressionEvaluatorContext createExpressionEvaluatorContext() {
        return new ExpressionEvaluatorContext(
                options.getExpressionEvaluatorFactory(),
                options.getExpressionNotationBegin(),
                options.getExpressionNotationEnd());
    }

    protected void preWrite() {
        transformer.setEvaluateFormulas(options.isRecalculateFormulasBeforeSaving());
        transformer.setFullFormulaRecalculationOnOpening(options.isRecalculateFormulasOnOpening());
		Consumer<String> action;
		switch (options.getKeepTemplateSheet()) {
		case DELETE:
			action = sheetName -> transformer.deleteSheet(sheetName);
			break;
		case HIDE:
			action = sheetName -> transformer.setHidden(sheetName, true);
			break;
		default:
			return;
		}
		getSheetsNameOfMultiSheetTemplate(areas).stream().forEach(action);
		options.getPreWriteActions().forEach(preWriteAction -> preWriteAction.preWrite(transformer, context));
    }

    /**
     * Return names of all multi sheet template
     *
     * @param areaList list of area
     * @return string array
     */
    public static List<String> getSheetsNameOfMultiSheetTemplate(List<Area> areaList) {
        List<String> templateSheetsName = new ArrayList<>();
        for (Area xlsArea : areaList) {
            for (Command command : xlsArea.findCommandByName("each")) {
                boolean isAreaHasMultiSheetAttribute = ((EachCommand) command).getMultisheet() != null && !((EachCommand) command).getMultisheet().isEmpty();
                if (isAreaHasMultiSheetAttribute) {
                    templateSheetsName.add(xlsArea.getAreaRef().getSheetName());
                    break;
                }
            }
        }
        return templateSheetsName;
    }

    protected void write() throws IOException {
        transformer.write();
    }
}
