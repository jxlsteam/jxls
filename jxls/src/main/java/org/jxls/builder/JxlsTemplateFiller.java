package org.jxls.builder;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;

import org.jxls.area.Area;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.command.Command;
import org.jxls.command.EachCommand;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.transform.TransformationConfig;
import org.jxls.transform.Transformer;

public class JxlsTemplateFiller {
    protected final JxlsOptions options;
    protected final InputStream template;
    protected Transformer transformer;
    protected List<Area> areas;
	private final Map<String, Class<? extends Command>> rem = new HashMap<>();

    protected JxlsTemplateFiller(JxlsOptions options, InputStream template) {
        this.options = options;
        this.template = template;
    }

    public void fill(Map<String, Object> data, JxlsOutput output) {
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
            areas = null;
            transformer = null;
            restoreCommands();
        }
    }

	private void installCommands() {
		options.getCommands().forEach((k, v) -> {
			rem.put(k, XlsCommentAreaBuilder.getCommandClass(k));
			XlsCommentAreaBuilder.addCommandMapping(k, v); // for the future we don't want it static
		});
	}

	private void restoreCommands() {
		rem.forEach((k, v) -> {
			if (v == null) {
				XlsCommentAreaBuilder.removeCommandMapping(k);
			} else {
				XlsCommentAreaBuilder.addCommandMapping(k, v);
			}
		});
		rem.clear();
	}

    protected void createTransformer(OutputStream outputStream) {
        transformer = options.getTransformerFactory().create(template, outputStream, options.getStreaming(), options.getLogger());
    }

    protected void configureTransformer() {
    	transformer.setIgnoreColumnProps(options.isIgnoreColumnProps());
    	transformer.setIgnoreRowProps(options.isIgnoreRowProps());
        TransformationConfig tc = transformer.getTransformationConfig();
        tc.buildExpressionNotation(options.getExpressionNotationBegin(), options.getExpressionNotationEnd());
        tc.setExpressionEvaluatorFactory(options.getExpressionEvaluatorFactory());
        options.getNeedsExpressionEvaluatorList().forEach(ee -> ee.setExpressionEvaluator(tc.getExpressionEvaluator()));
    }

    /**
     * Implementation must set areas variable.
     * @param data -
     */
    protected void processAreas(Map<String, Object> data) {
        areas = options.getAreaBuilder().build(transformer, options.isClearTemplateCells());
        Context context = new Context(data);
        for (Area area : areas) {
            area.applyAt(new CellRef(area.getStartCellRef().getCellName()), context);
        }
        if (options.getFormulaProcessor() != null) {
            for (Area area : areas) {
                area.setFormulaProcessor(options.getFormulaProcessor());
                area.processFormulas();
            }
        }
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
