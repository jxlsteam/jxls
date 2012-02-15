package com.jxls.writer.area;

import com.jxls.writer.command.Command;
import com.jxls.writer.common.Context;
import com.jxls.writer.common.*;
import com.jxls.writer.transform.Transformer;
import com.jxls.writer.util.CellRefUtil;
import com.jxls.writer.util.Util;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author Leonid Vysochyn
 * Date: 1/16/12 6:44 PM
 */
public class XlsArea implements Area {
    static Logger logger = LoggerFactory.getLogger(XlsArea.class);

    public static final XlsArea EMPTY_AREA = new XlsArea(new CellRef(null,0, 0), Size.ZERO_SIZE);

    List<CommandData> commandDataList = new ArrayList<CommandData>();
    Transformer transformer;
    
    CellRange cellRange;
    
    CellRef startCellRef;
    Size size;

    public XlsArea(AreaRef areaRef, Transformer transformer){
        CellRef startCell = areaRef.getFirstCellRef();
        CellRef endCell = areaRef.getLastCellRef();
        this.startCellRef = startCell;
        this.size = new Size(endCell.getCol() - startCell.getCol() + 1, endCell.getRow() - startCell.getRow() + 1);
        this.transformer = transformer;
    }
    
    public XlsArea(String areaRef, Transformer transformer){
        this(new AreaRef(areaRef), transformer);
    }
    
    public XlsArea(CellRef startCell, CellRef endCell, Transformer transformer){
        this(new AreaRef(startCell, endCell), transformer);
    }

    public XlsArea(CellRef startCellRef, Size size, List<CommandData> commandDataList, Transformer transformer) {
        this.startCellRef = startCellRef;
        this.size = size;
        this.commandDataList = commandDataList != null ? commandDataList : new ArrayList<CommandData>();
        this.transformer = transformer;
    }

    public XlsArea(CellRef startCellRef, Size size) {
        this(startCellRef, size, null, null);
    }

    public XlsArea(CellRef startCellRef, Size size, Transformer transformer) {
        this(startCellRef, size, null, transformer);
    }

    public void addCommand(CellRef cellRef, Size size, Command command){
        commandDataList.add(new CommandData(cellRef, size, command));
    }
    
    public void addCommand(AreaRef areaRef, Command command){
        commandDataList.add(new CommandData(areaRef, command));
    }

    public void addCommand(String areaRef, Command command){
        commandDataList.add(new CommandData(areaRef, command));
    }

    public List<CommandData> getCommandDataList() {
        return commandDataList;
    }

    public Transformer getTransformer() {
        return transformer;
    }

    public void setTransformer(Transformer transformer) {
        this.transformer = transformer;
    }
    
    private void createCellRange(){
        cellRange = new CellRange(startCellRef, size.getWidth(), size.getHeight());
        for(CommandData commandData: commandDataList){
            cellRange.excludeCells(commandData.getStartCellRef().getCol() - startCellRef.getCol(), commandData.getStartCellRef().getCol() - startCellRef.getCol() + commandData.getSize().getWidth()-1,
                    commandData.getStartCellRef().getRow() - startCellRef.getRow(), commandData.getStartCellRef().getRow() - startCellRef.getRow() + commandData.getSize().getHeight()-1);
        }
    }

    public Size applyAt(CellRef cellRef, Context context) {
        logger.debug("Applying XlsArea at {} with {}", cellRef, context);
        int widthDelta = 0;
        int heightDelta = 0;
        createCellRange();
        for (int i = 0; i < commandDataList.size(); i++) {
            cellRange.resetChangeMatrix();
            CommandData commandData = commandDataList.get(i);
            CellRef newCell = new CellRef(cellRef.getSheetName(), commandData.getStartCellRef().getRow() - startCellRef.getRow() + cellRef.getRow(), commandData.getStartCellRef().getCol() - startCellRef.getCol() + cellRef.getCol());
            Size initialSize = commandData.getSize();
            Size newSize = commandData.getCommand().applyAt(newCell, context);
            int widthChange = newSize.getWidth() - initialSize.getWidth();
            int heightChange = newSize.getHeight() - initialSize.getHeight();
            if( widthChange != 0 || heightChange != 0){
                widthDelta += widthChange;
                heightDelta += heightChange;
                if( widthChange != 0 ){
                    cellRange.shiftCellsWithRowBlock(commandData.getStartCellRef().getRow() - startCellRef.getRow(),
                            commandData.getStartCellRef().getRow() - startCellRef.getRow() + commandData.getSize().getHeight(),
                            commandData.getStartCellRef().getCol() - startCellRef.getCol() + initialSize.getWidth(), widthChange);
                }
                if( heightChange != 0 ){
                    cellRange.shiftCellsWithColBlock(commandData.getStartCellRef().getCol() - startCellRef.getCol(),
                            commandData.getStartCellRef().getCol() - startCellRef.getCol() + newSize.getWidth()-1, commandData.getStartCellRef().getRow() - startCellRef.getRow() + commandData.getSize().getHeight()-1, heightChange);
                }
                for (int j = i + 1; j < commandDataList.size(); j++) {
                    CommandData data = commandDataList.get(j);
                    int newRow = data.getStartCellRef().getRow() - startCellRef.getRow() + cellRef.getRow();
                    int newCol = data.getStartCellRef().getCol() - startCellRef.getCol() + cellRef.getCol();
                    if(newRow > newCell.getRow() && ((newCol >= newCell.getCol() && newCol <= newCell.getCol() + newSize.getWidth()) ||
                            (newCol + data.getSize().getWidth() >= newCell.getCol() && newCol + data.getSize().getWidth() <= newCell.getCol() + newSize.getWidth()) ||
                            (newCell.getCol() >= newCol && newCell.getCol() <= newCol + data.getSize().getWidth() )
                    )){
                        cellRange.shiftCellsWithColBlock(data.getStartCellRef().getCol() - startCellRef.getCol(),
                                data.getStartCellRef().getCol() - startCellRef.getCol() + data.getSize().getWidth()-1, data.getStartCellRef().getRow() - startCellRef.getRow() + data.getSize().getHeight()-1, heightChange);
                        data.setStartCellRef(new CellRef(data.getStartCellRef().getSheetName(), data.getStartCellRef().getRow() + heightChange, data.getStartCellRef().getCol()));
                    }else
                    if( newCol > newCell.getCol() && ( (newRow >= newCell.getRow() && newRow <= newCell.getRow() + newSize.getHeight()) ||
                   ( newRow + data.getSize().getHeight() >= newCell.getRow() && newRow + data.getSize().getHeight() <= newCell.getRow() + newSize.getHeight()) ||
                    newCell.getRow() >= newRow && newCell.getRow() <= newRow + data.getSize().getHeight()) ){
                        cellRange.shiftCellsWithRowBlock(data.getStartCellRef().getRow() - startCellRef.getRow(),
                                data.getStartCellRef().getRow() - startCellRef.getRow() + data.getSize().getHeight()-1,
                                data.getStartCellRef().getCol() - startCellRef.getCol() + initialSize.getWidth(), widthChange);
                        data.setStartCellRef(new CellRef(data.getStartCellRef().getSheetName(), data.getStartCellRef().getRow(), data.getStartCellRef().getCol() + widthChange));
                    }
                }
            }
        }
        transformStaticCells(cellRef, context);
        return new Size(size.getWidth() + widthDelta, size.getHeight() + heightDelta);
    }

    public void clearCells() {
        String sheetName = startCellRef.getSheetName();
        int startRow = startCellRef.getRow();
        int startCol = startCellRef.getCol();
        for(int row = 0; row < size.getHeight(); row++){
            for(int col = 0; col < size.getWidth(); col++){
                CellRef cellRef = new CellRef(sheetName, startRow + row, startCol + col);
                transformer.clearCell(cellRef);
            }
        }
    }

    private void transformStaticCells(CellRef cellRef, Context context) {
        String sheetName = startCellRef.getSheetName();
        int startRow = startCellRef.getRow();
        int startCol = startCellRef.getCol();
        for(int col = 0; col < size.getWidth(); col++){
            for(int row = 0; row < size.getHeight(); row++){
                if( !cellRange.isExcluded(row, col) ){
                    CellRef relativeCell = cellRange.getCell(row, col);
                    CellRef srcCell = new CellRef(sheetName, startRow + row, startCol + col);
                    CellRef targetCell = new CellRef(cellRef.getSheetName(), relativeCell.getRow() + cellRef.getRow(), relativeCell.getCol() + cellRef.getCol());
                    transformer.transform(srcCell, targetCell, context);
                }
            }
        }
    }

    public CellRef getStartCellRef() {
        return startCellRef;
    }

    public Size getSize() {
        return size;
    }


    public void processFormulas() {
        Set<CellData> formulaCells = transformer.getFormulaCells();
        for (CellData formulaCellData : formulaCells) {
            List<String> formulaCellRefs = Util.getFormulaCellRefs(formulaCellData.getFormula());
            List<String> jointedCellRefs = Util.getJointedCellRefs(formulaCellData.getFormula());
            List<CellRef> targetFormulaCells = transformer.getTargetCellRef(formulaCellData.getCellRef());
            Map<CellRef, List<CellRef>> targetCellRefMap = new HashMap<CellRef, List<CellRef>>();
            Map<String, List<CellRef>> jointedCellRefMap = new HashMap<String, List<CellRef>>();
            for (String cellRef : formulaCellRefs) {
                CellRef pos = new CellRef(cellRef);
                if(pos.getSheetName() == null ){
                    pos.setSheetName( formulaCellData.getSheetName() );
                    pos.setIgnoreSheetNameInFormat(true);
                }
                List<CellRef> targetCellDataList = transformer.getTargetCellRef(pos);
                targetCellRefMap.put(pos, targetCellDataList);
            }
            for (String jointedCellRef : jointedCellRefs) {
                List<String> nestedCellRefs = Util.getCellRefsFromJointedCellRef(jointedCellRef);
                List<CellRef> jointedCellRefList = new ArrayList<CellRef>();
                for (String cellRef : nestedCellRefs) {
                    CellRef pos = new CellRef(cellRef);
                    if(pos.getSheetName() == null ){
                        pos.setSheetName(formulaCellData.getSheetName());
                        pos.setIgnoreSheetNameInFormat(true);
                    }
                    List<CellRef> targetCellDataList = transformer.getTargetCellRef(pos);

                    jointedCellRefList.addAll(targetCellDataList);
                }
                jointedCellRefMap.put(jointedCellRef, jointedCellRefList);
            }
            for (int i = 0; i < targetFormulaCells.size(); i++) {
                CellRef targetFormulaCellRef = targetFormulaCells.get(i);
                String targetFormulaString = formulaCellData.getFormula();
                for (Map.Entry<CellRef, List<CellRef>> cellRefEntry : targetCellRefMap.entrySet()) {
                    List<CellRef> targetCells = cellRefEntry.getValue();
                    if( targetCells.isEmpty() ) continue;
                    if( targetCells.size() == targetFormulaCells.size() ){
                        CellRef targetCellRefCellRef = targetCells.get(i);
                        targetFormulaString = targetFormulaString.replaceAll(Util.regexJointedLookBehind + sheetNameRegex(cellRefEntry) + Pattern.quote(cellRefEntry.getKey().getCellName()), Matcher.quoteReplacement( targetCellRefCellRef.getCellName() ));
                    }else{
                        List< List<CellRef> > rangeList = Util.groupByRanges(targetCells, targetFormulaCells.size());
                        if( rangeList.size() == targetFormulaCells.size() ){
                            List<CellRef> range = rangeList.get(i);
                            String replacementString = Util.createTargetCellRef( range );
                            targetFormulaString = targetFormulaString.replaceAll(Util.regexJointedLookBehind + sheetNameRegex(cellRefEntry) + Pattern.quote(cellRefEntry.getKey().getCellName()), Matcher.quoteReplacement(replacementString));
                        }else{
                            targetFormulaString = targetFormulaString.replaceAll(Util.regexJointedLookBehind + sheetNameRegex(cellRefEntry) + Pattern.quote(cellRefEntry.getKey().getCellName()), Matcher.quoteReplacement(Util.createTargetCellRef(targetCells)));
                        }
                    }
                }
                for (Map.Entry<String, List<CellRef>> jointedCellRefEntry : jointedCellRefMap.entrySet()) {
                    List<CellRef> targetCellRefList = jointedCellRefEntry.getValue();
                    if( targetCellRefList.isEmpty() ) continue;
                    List< List<CellRef> > rangeList = Util.groupByRanges(targetCellRefList, targetFormulaCells.size());
                    if( rangeList.size() == targetFormulaCells.size() ){
                        List<CellRef> range = rangeList.get(i);
                        String replacementString = Util.createTargetCellRef(range);
                        targetFormulaString = targetFormulaString.replaceAll(Pattern.quote(jointedCellRefEntry.getKey()), replacementString);
                    }else{
                        targetFormulaString = targetFormulaString.replaceAll( Pattern.quote(jointedCellRefEntry.getKey()), Util.createTargetCellRef(targetCellRefList));
                    }
                }
                String sheetNameReplacementRegex = targetFormulaCellRef.getFormattedSheetName() + CellRefUtil.SHEET_NAME_DELIMITER;
                targetFormulaString = targetFormulaString.replaceAll(sheetNameReplacementRegex, "");
                transformer.setFormula(new CellRef(targetFormulaCellRef.getSheetName(), targetFormulaCellRef.getRow(), targetFormulaCellRef.getCol()), targetFormulaString);
            }
        }
    }

    private String sheetNameRegex(Map.Entry<CellRef, List<CellRef>> cellRefEntry) {
        return (cellRefEntry.getKey().isIgnoreSheetNameInFormat()?"(?<!!)":"");
    }


}
