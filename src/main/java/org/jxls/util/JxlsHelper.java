package org.jxls.util;

import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.command.GridCommand;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.formula.FastFormulaProcessor;
import org.jxls.formula.StandardFormulaProcessor;
import org.jxls.template.SimpleExporter;
import org.jxls.transform.Transformer;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Collection;
import java.util.List;

/**
 * Common utilities
 */
public class JxlsHelper {
    private boolean hideTemplateSheet = false;
    private boolean deleteTemplateSheet = true;
    private boolean processFormulas = true;
    private boolean useFastFormulaProcessor = true;
    private String expressionNotationBegin;
    private String expressionNotationEnd;
    private SimpleExporter simpleExporter = new SimpleExporter();

    private AreaBuilder areaBuilder = new XlsCommentAreaBuilder();

    public static JxlsHelper getInstance(){
        return new JxlsHelper();
    }

    private JxlsHelper() {
    }

    public AreaBuilder getAreaBuilder() {
        return areaBuilder;
    }

    public JxlsHelper setAreaBuilder(AreaBuilder areaBuilder) {
        this.areaBuilder = areaBuilder;
        return this;
    }

    public boolean isProcessFormulas() {
        return processFormulas;
    }

    public JxlsHelper setProcessFormulas(boolean processFormulas) {
        this.processFormulas = processFormulas;
        return this;
    }

    public boolean isHideTemplateSheet() {
        return hideTemplateSheet;
    }

    public JxlsHelper setHideTemplateSheet(boolean hideTemplateSheet) {
        this.hideTemplateSheet = hideTemplateSheet;
        return this;
    }

    public boolean isDeleteTemplateSheet() {
        return deleteTemplateSheet;
    }

    public JxlsHelper setDeleteTemplateSheet(boolean deleteTemplateSheet) {
        this.deleteTemplateSheet = deleteTemplateSheet;
        return this;
    }

    public boolean isUseFastFormulaProcessor() {
        return useFastFormulaProcessor;
    }

    public JxlsHelper setUseFastFormulaProcessor(boolean useFastFormulaProcessor) {
        this.useFastFormulaProcessor = useFastFormulaProcessor;
        return this;
    }

    public JxlsHelper buildExpressionNotation(String expressionNotationBegin, String expressionNotationEnd){
        this.expressionNotationBegin = expressionNotationBegin;
        this.expressionNotationEnd = expressionNotationEnd;
        return this;
    }

    public JxlsHelper processTemplate(InputStream templateStream, OutputStream targetStream, Context context) throws IOException {
        Transformer transformer = createTransformer(templateStream, targetStream);
        processTemplate(context, transformer);
        return this;
    }

    public void processTemplate(Context context, Transformer transformer) throws IOException {
        areaBuilder.setTransformer(transformer);
        List<Area> xlsAreaList = areaBuilder.build();
        for (Area xlsArea : xlsAreaList) {
            xlsArea.applyAt(
                    new CellRef(xlsArea.getStartCellRef().getCellName()), context);
            if( processFormulas ) {
                setFormulaProcessor(xlsArea);
                xlsArea.processFormulas();
            }
        }
        transformer.write();
    }

    private Area setFormulaProcessor(Area xlsArea) {
        if( useFastFormulaProcessor ){
            xlsArea.setFormulaProcessor(new FastFormulaProcessor());
        }else{
            xlsArea.setFormulaProcessor(new StandardFormulaProcessor());
        }
        return xlsArea;
    }

    public JxlsHelper processTemplateAtCell(InputStream templateStream, OutputStream targetStream, Context context, String targetCell) throws IOException {
        Transformer transformer = createTransformer(templateStream, targetStream);
        areaBuilder.setTransformer(transformer);
        List<Area> xlsAreaList = areaBuilder.build();
        if( xlsAreaList.isEmpty() ){
            throw new IllegalStateException("No XlsArea were detected for this processing");
        }
        Area firstArea = xlsAreaList.get(0);
        CellRef targetCellRef = new CellRef(targetCell);
        firstArea.applyAt(targetCellRef, context);
        if( processFormulas ){
            setFormulaProcessor(firstArea);
            firstArea.processFormulas();
        }
        String sourceSheetName = firstArea.getStartCellRef().getSheetName();
        if( !sourceSheetName.equalsIgnoreCase(targetCellRef.getSheetName())){
            if( hideTemplateSheet ){
                transformer.setHidden(sourceSheetName, true);
            }
            if( deleteTemplateSheet ){
                transformer.deleteSheet(sourceSheetName);
            }
        }
        transformer.write();
        return this;
    }

    public JxlsHelper processGridTemplate(InputStream templateStream, OutputStream targetStream, Context context, String objectProps) throws IOException {
        Transformer transformer = createTransformer(templateStream, targetStream);
        areaBuilder.setTransformer(transformer);
        List<Area> xlsAreaList = areaBuilder.build();
        for (Area xlsArea : xlsAreaList) {
            GridCommand gridCommand = (GridCommand) xlsArea.getCommandDataList().get(0).getCommand();
            gridCommand.setProps(objectProps);
            setFormulaProcessor(xlsArea);
            xlsArea.applyAt(
                    new CellRef(xlsArea.getStartCellRef().getCellName()), context);
            if( processFormulas ) {
                xlsArea.processFormulas();
            }
        }
        transformer.write();
        return this;
    }

    public void processGridTemplateAtCell(InputStream templateStream, OutputStream targetStream, Context context, String objectProps, String targetCell) throws IOException {
        Transformer transformer = createTransformer(templateStream, targetStream);
        areaBuilder.setTransformer(transformer);
        List<Area> xlsAreaList = areaBuilder.build();
        Area firstArea = xlsAreaList.get(0);
        CellRef targetCellRef = new CellRef(targetCell);
        GridCommand gridCommand = (GridCommand) firstArea.getCommandDataList().get(0).getCommand();
        gridCommand.setProps(objectProps);
        firstArea.applyAt(targetCellRef, context);
        if( processFormulas ){
            setFormulaProcessor(firstArea);
            firstArea.processFormulas();
        }
        String sourceSheetName = firstArea.getStartCellRef().getSheetName();
        if( !sourceSheetName.equalsIgnoreCase(targetCellRef.getSheetName())){
            if( hideTemplateSheet ){
                transformer.setHidden(sourceSheetName, true);
            }
            if( deleteTemplateSheet ){
                transformer.deleteSheet(sourceSheetName);
            }
        }
        transformer.write();
    }

    public JxlsHelper registerGridTemplate(InputStream inputStream) throws IOException {
        simpleExporter.registerGridTemplate(inputStream);
        return this;
    }

    public JxlsHelper gridExport(Collection headers, Collection dataObjects, String objectProps, OutputStream outputStream){
        simpleExporter.gridExport(headers, dataObjects, objectProps, outputStream);
        return this;
    }

    public Transformer createTransformer(InputStream templateStream, OutputStream targetStream) {
        Transformer transformer = TransformerFactory.createTransformer(templateStream, targetStream);
        if( transformer == null ){
            throw new IllegalStateException("Cannot load XLS transformer. Please make sure a Transformer implementation is in classpath");
        }
        if( expressionNotationBegin != null && expressionNotationEnd != null){
            transformer.getTransformationConfig().buildExpressionNotation(expressionNotationBegin, expressionNotationEnd);
        }
        return transformer;
    }

}
