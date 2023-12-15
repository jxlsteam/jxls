package org.jxls.builder;

import static org.jxls.util.Util.getSheetsNameOfMultiSheetTemplate;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;

import org.jxls.area.Area;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.command.Command;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.ExceptionHandler;
import org.jxls.expression.ExpressionEvaluatorFactory;
import org.jxls.formula.FormulaProcessor;
import org.jxls.transform.JxlsTransformerFactory;
import org.jxls.transform.Transformer;

public class JxlsTemplateFiller {
    protected final ExpressionEvaluatorFactory expressionEvaluatorFactory;
    protected final String expressionNotationBegin;
    protected final String expressionNotationEnd;
    protected final ExceptionHandler exceptionHandler;
    protected final FormulaProcessor formulaProcessor;
    protected final boolean recalculateFormulasBeforeSaving;
    protected final boolean recalculateFormulasOnOpening;
    protected final KeepTemplateSheet keepTemplateSheet;
    protected final AreaBuilder areaBuilder;
    protected final Map<String, Class<? extends Command>> commands;
    protected final boolean clearTemplateCells;
    protected final JxlsTransformerFactory transformerFactory;
    protected final JxlsStreaming streaming;
    protected final InputStream template;
    protected Transformer transformer;
    protected List<Area> areas;
	private final Map<String, Class<? extends Command>> rem = new HashMap<>();

    protected JxlsTemplateFiller(ExpressionEvaluatorFactory expressionEvaluatorFactory,
            String expressionNotationBegin, String expressionNotationEnd, ExceptionHandler exceptionHandler, //
            FormulaProcessor formulaProcessor, boolean recalculateFormulasBeforeSaving, boolean recalculateFormulasOnOpening, //
            KeepTemplateSheet keepTemplateSheet, AreaBuilder areaBuilder, Map<String, Class<? extends Command>> commands,
            boolean clearTemplateCells, JxlsTransformerFactory transformerFactory, JxlsStreaming streaming, //
            InputStream template) {
        this.expressionEvaluatorFactory = expressionEvaluatorFactory;
        this.expressionNotationBegin = expressionNotationBegin;
        this.expressionNotationEnd = expressionNotationEnd;
        this.exceptionHandler = exceptionHandler;
        this.recalculateFormulasBeforeSaving = recalculateFormulasBeforeSaving;
        this.recalculateFormulasOnOpening = recalculateFormulasOnOpening;
        this.formulaProcessor = formulaProcessor;
        this.keepTemplateSheet = keepTemplateSheet;
        this.areaBuilder = areaBuilder;
        this.commands = commands;
        this.clearTemplateCells = clearTemplateCells;
        this.transformerFactory = transformerFactory;
        this.streaming = streaming;
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
		commands.forEach((k, v) -> {
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
        transformer = transformerFactory.create(template, outputStream, streaming);
    }

    protected void configureTransformer() {
		if (exceptionHandler != null) {
			transformer.setExceptionHandler(exceptionHandler);
		}
        transformer.getTransformationConfig().buildExpressionNotation(expressionNotationBegin, expressionNotationEnd);
        transformer.getTransformationConfig().setExpressionEvaluatorFactory(expressionEvaluatorFactory);
    }

    protected void processAreas(Map<String, Object> data) {
        areas = areaBuilder.build(transformer, clearTemplateCells);
        Context context = new Context(data);
        for (Area area : areas) {
            area.applyAt(new CellRef(area.getStartCellRef().getCellName()), context);
        }
        if (formulaProcessor != null) {
            for (Area area : areas) {
                area.setFormulaProcessor(formulaProcessor);
                area.processFormulas();
            }
        }
    }

    protected void preWrite() {
        transformer.setEvaluateFormulas(recalculateFormulasBeforeSaving);
        transformer.setFullFormulaRecalculationOnOpening(recalculateFormulasOnOpening);
		Consumer<String> action;
		switch (keepTemplateSheet) {
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

    protected void write() throws IOException {
        transformer.write();
    }
}