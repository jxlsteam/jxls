package org.jxls.area;

import org.jxls.command.Command;
import org.jxls.common.*;
import org.jxls.common.cellshift.AdjacentCellShiftStrategy;
import org.jxls.common.cellshift.CellShiftStrategy;
import org.jxls.common.cellshift.InnerCellShiftStrategy;
import org.jxls.formula.FastFormulaProcessor;
import org.jxls.formula.FormulaProcessor;
import org.jxls.transform.Transformer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.*;

/**
 * Core implementation of {@link Area} interface
 * @author Leonid Vysochyn
 * Date: 1/16/12
 */
public class XlsArea implements Area {
    static Logger logger = LoggerFactory.getLogger(XlsArea.class);

    public static final XlsArea EMPTY_AREA = new XlsArea(new CellRef(null,0, 0), Size.ZERO_SIZE);

    List<CommandData> commandDataList = new ArrayList<CommandData>();
    Transformer transformer;
    
    CellRange cellRange;
    
    CellRef startCellRef;
    Size size;
    List<AreaListener> areaListeners = new ArrayList<AreaListener>();

    private boolean cellsCleared = false;

    private FormulaProcessor formulaProcessor = new FastFormulaProcessor();
    // default cell shift strategy
    private CellShiftStrategy cellShiftStrategy = new InnerCellShiftStrategy();

    private final CellShiftStrategy innerCellShiftStrategy = new InnerCellShiftStrategy();
    private final CellShiftStrategy adjacentCellShiftStrategy = new AdjacentCellShiftStrategy();

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

    @Override
    public CellShiftStrategy getCellShiftStrategy() {
        return cellShiftStrategy;
    }

    @Override
    public void setCellShiftStrategy(CellShiftStrategy cellShiftStrategy) {
        this.cellShiftStrategy = cellShiftStrategy;
    }

    @Override
    public FormulaProcessor getFormulaProcessor() {
        return formulaProcessor;
    }

    @Override
    public void setFormulaProcessor(FormulaProcessor formulaProcessor) {
        this.formulaProcessor = formulaProcessor;
    }

    public void addCommand(AreaRef areaRef, Command command){
        AreaRef thisAreaRef = new AreaRef(startCellRef, size);
        if( !thisAreaRef.contains(areaRef) ){
            throw new IllegalArgumentException("Cannot add command '" + command.getName() + "' to area " + thisAreaRef + " at " + areaRef);
        }
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
            CellRef startCellRef = commandData.getSourceStartCellRef();
            Size size = commandData.getSourceSize();
            cellRange.excludeCells(startCellRef.getCol() - this.startCellRef.getCol(), startCellRef.getCol() - this.startCellRef.getCol() + size.getWidth()-1,
                    startCellRef.getRow() - this.startCellRef.getRow(), startCellRef.getRow() - this.startCellRef.getRow() + size.getHeight()-1);
        }
    }


    public Size applyAt(CellRef cellRef, Context context) {
        logger.debug("Applying XlsArea at {} with {}", cellRef, context);
        fireBeforeApplyEvent(cellRef, context);
        createCellRange();
        int topStaticAreaLastRow = transformTopStaticArea(cellRef, context);
        for (int i = 0; i < commandDataList.size(); i++) {
            cellRange.resetChangeMatrix();
            CommandData commandData = commandDataList.get(i);
            String shiftMode = commandData.getCommand().getShiftMode();
            CellShiftStrategy commandCellShiftStrategy = detectCellShiftStrategy(shiftMode);
            cellRange.setCellShiftStrategy(commandCellShiftStrategy);
            CellRef commandStartCellRef = commandData.getStartCellRef();
            Size commandInitialSize = commandData.getSize();
            CellRef newCell = new CellRef(cellRef.getSheetName(), commandStartCellRef.getRow() - this.startCellRef.getRow() + cellRef.getRow(), commandStartCellRef.getCol() - this.startCellRef.getCol() + cellRef.getCol());
            Size commandNewSize = commandData.getCommand().applyAt(newCell, context);
            int widthChange = commandNewSize.getWidth() - commandInitialSize.getWidth();
            int heightChange = commandNewSize.getHeight() - commandInitialSize.getHeight();
            if( widthChange != 0 || heightChange != 0){
                if( widthChange != 0 ){
                    cellRange.shiftCellsWithRowBlock(commandStartCellRef.getRow() - this.startCellRef.getRow(),
                            commandStartCellRef.getRow() - this.startCellRef.getRow() + commandData.getSize().getHeight(),
                            commandStartCellRef.getCol() - this.startCellRef.getCol() + commandInitialSize.getWidth()-1, widthChange, true);
                }
                if( heightChange != 0 ){
                    cellRange.shiftCellsWithColBlock(commandStartCellRef.getCol() - this.startCellRef.getCol(),
                            commandStartCellRef.getCol() - this.startCellRef.getCol() + commandNewSize.getWidth()-1, commandStartCellRef.getRow() - this.startCellRef.getRow() + commandData.getSize().getHeight()-1, heightChange, true);
                }
                for (int j = i + 1; j < commandDataList.size(); j++) {
                    CommandData data = commandDataList.get(j);
                    CellRef commandDataStartCellRef = data.getStartCellRef();
                    int newRow = commandDataStartCellRef.getRow() - this.startCellRef.getRow() + cellRef.getRow();
                    int newCol = commandDataStartCellRef.getCol() - this.startCellRef.getCol() + cellRef.getCol();
                    Size commandDataSize = data.getSize();
                    if(newRow > newCell.getRow() && ((newCol >= newCell.getCol() && newCol <= newCell.getCol() + commandNewSize.getWidth()) ||
                            (newCol + commandDataSize.getWidth() >= newCell.getCol() && newCol + commandDataSize.getWidth() <= newCell.getCol() + commandNewSize.getWidth()) ||
                            (newCell.getCol() >= newCol && newCell.getCol() <= newCol + commandDataSize.getWidth() )
                    )){
                        cellRange.shiftCellsWithColBlock(commandDataStartCellRef.getCol() - this.startCellRef.getCol(),
                                commandDataStartCellRef.getCol() - this.startCellRef.getCol() + commandDataSize.getWidth()-1, commandDataStartCellRef.getRow() - this.startCellRef.getRow() + commandDataSize.getHeight()-1, heightChange, false);
                        data.setStartCellRef(new CellRef(commandDataStartCellRef.getSheetName(), commandDataStartCellRef.getRow() + heightChange, commandDataStartCellRef.getCol()));
                    }else
                    if( newCol > newCell.getCol() && ( (newRow >= newCell.getRow() && newRow <= newCell.getRow() + commandNewSize.getHeight()) ||
                   ( newRow + commandDataSize.getHeight() >= newCell.getRow() && newRow + commandDataSize.getHeight() <= newCell.getRow() + commandNewSize.getHeight()) ||
                    newCell.getRow() >= newRow && newCell.getRow() <= newRow + commandDataSize.getHeight()) ){
                        cellRange.shiftCellsWithRowBlock(commandDataStartCellRef.getRow() - this.startCellRef.getRow(),
                                commandDataStartCellRef.getRow() - this.startCellRef.getRow() + commandDataSize.getHeight()-1,
                                commandDataStartCellRef.getCol() - this.startCellRef.getCol() + commandInitialSize.getWidth(), widthChange, false);
                        data.setStartCellRef(new CellRef(commandDataStartCellRef.getSheetName(), commandDataStartCellRef.getRow(), commandDataStartCellRef.getCol() + widthChange));
                    }
                }
            }
        }
        transformStaticCells(cellRef, context, topStaticAreaLastRow + 1);
        fireAfterApplyEvent(cellRef, context);
        Size finalSize = new Size(cellRange.calculateWidth(), cellRange.calculateHeight());
        AreaRef newAreaRef = new AreaRef(cellRef, finalSize);
        updateCellDataFinalAreaForFormulaCells(newAreaRef);
        for(CommandData commandData: commandDataList){
            commandData.resetStartCellAndSize();
        }
        return finalSize;
    }

    private CellShiftStrategy detectCellShiftStrategy(String shiftMode) {
        if( shiftMode != null && Command.ADJACENT_SHIFT_MODE.equalsIgnoreCase(shiftMode)){
            return adjacentCellShiftStrategy;
        }else{
            return innerCellShiftStrategy;
        }
    }

    private void updateCellDataFinalAreaForFormulaCells(AreaRef newAreaRef) {
        String sheetName = startCellRef.getSheetName();
        int offsetRow = startCellRef.getRow();
        int startCol = startCellRef.getCol();
        for(int col = 0; col < size.getWidth(); col++){
            for(int row = 0; row < size.getHeight(); row++){
                if( !cellRange.isExcluded(row, col) ){
                    CellRef srcCell = new CellRef(sheetName, offsetRow + row, startCol + col);
                    CellData cellData = transformer.getCellData(srcCell);
                    if( cellData != null && cellData.isFormulaCell() ){
                        cellData.addTargetParentAreaRef( newAreaRef );
                    }
                }
            }
        }
    }

    private int transformTopStaticArea(CellRef cellRef, Context context) {
        String sheetName = startCellRef.getSheetName();
        int startRow = startCellRef.getRow();
        int startCol = startCellRef.getCol();
        int topStaticAreaLastRow = findRelativeTopCommandRow() - 1;
        for(int col = 0; col < size.getWidth(); col++){
            for(int row = 0; row <= topStaticAreaLastRow; row++){
                if( !cellRange.isExcluded(row, col) ){
                    CellRef relativeCell = cellRange.getCell(row, col);
                    CellRef srcCell = new CellRef(sheetName, startRow + row, startCol + col);
                    CellRef targetCell = new CellRef(cellRef.getSheetName(), relativeCell.getRow() + cellRef.getRow(), relativeCell.getCol() + cellRef.getCol());
                    fireBeforeTransformCell(srcCell, targetCell, context);
                    try{
                        updateCellDataArea(srcCell, targetCell, context);
                        transformer.transform(srcCell, targetCell, context);
                    }catch(Exception e){
                        logger.error("Failed to transform " + srcCell + " into " + targetCell, e);
                    }
                    fireAfterTransformCell(srcCell, targetCell, context);
                }
            }
        }
        return topStaticAreaLastRow;
    }

    private int findRelativeTopCommandRow() {
        int topCommandRow = startCellRef.getRow() + size.getHeight() - 1;
        for(CommandData data : commandDataList){
            topCommandRow = Math.min(data.getStartCellRef().getRow(), topCommandRow);
        }
        return topCommandRow - startCellRef.getRow();
    }

    private void fireBeforeApplyEvent(CellRef cellRef, Context context) {
        for (AreaListener areaListener : areaListeners) {
            areaListener.beforeApplyAtCell(cellRef, context);
        }
    }

    private void fireAfterApplyEvent(CellRef cellRef, Context context) {
        for (AreaListener areaListener : areaListeners) {
            areaListener.afterApplyAtCell(cellRef, context);
        }
    }


    public void clearCells() {
        if( cellsCleared ) return;
        String sheetName = startCellRef.getSheetName();
        int startRow = startCellRef.getRow();
        int startCol = startCellRef.getCol();
        for(int row = 0; row < size.getHeight(); row++){
            for(int col = 0; col < size.getWidth(); col++){
                CellRef cellRef = new CellRef(sheetName, startRow + row, startCol + col);
                transformer.clearCell(cellRef);
            }
        }
        transformer.resetArea(getAreaRef());
        cellsCleared = true;
    }

    private void transformStaticCells(CellRef cellRef, Context context, int relativeStartRow) {
        String sheetName = startCellRef.getSheetName();
        int offsetRow = startCellRef.getRow();
        int startCol = startCellRef.getCol();
        for(int col = 0; col < size.getWidth(); col++){
            for(int row = relativeStartRow; row < size.getHeight(); row++){
                if( !cellRange.isExcluded(row, col) ){
                    CellRef relativeCell = cellRange.getCell(row, col);
                    CellRef srcCell = new CellRef(sheetName, offsetRow + row, startCol + col);
                    CellRef targetCell = new CellRef(cellRef.getSheetName(), relativeCell.getRow() + cellRef.getRow(), relativeCell.getCol() + cellRef.getCol());
                    fireBeforeTransformCell(srcCell, targetCell, context);
                    try{
                        updateCellDataArea(srcCell, targetCell, context);
                        transformer.transform(srcCell, targetCell, context);
                    }catch(Exception e){
                        logger.error("Failed to transform " + srcCell + " into " + targetCell, e);
                    }
                    fireAfterTransformCell(srcCell, targetCell, context);
                }
            }
        }
    }

    private void updateCellDataArea(CellRef srcCell, CellRef targetCell, Context context) {
        Context.Config config = context.getConfig();
        if( !config.isFormulaProcessingRequired() ) return;
        CellData cellData = transformer.getCellData(srcCell);
        if( cellData != null ) {
            cellData.setArea(this);
            cellData.addTargetPos(targetCell);
        }
    }

    private void fireBeforeTransformCell(CellRef srcCell, CellRef targetCell, Context context) {
        for (AreaListener areaListener : areaListeners) {
            areaListener.beforeTransformCell(srcCell, targetCell, context);
        }
    }

    private void fireAfterTransformCell(CellRef srcCell, CellRef targetCell, Context context) {
        for (AreaListener areaListener : areaListeners) {
            areaListener.afterTransformCell(srcCell, targetCell, context);
        }
    }

    public CellRef getStartCellRef() {
        return startCellRef;
    }

    public Size getSize() {
        return size;
    }
    
    public AreaRef getAreaRef(){
        return new AreaRef(startCellRef, size);
    }


    public void processFormulas() {
        formulaProcessor.processAreaFormulas(transformer);
    }

    public void addAreaListener(AreaListener listener) {
        areaListeners.add(listener);
    }

    public List<AreaListener> getAreaListeners() {
        return areaListeners;
    }

    public List<Command> findCommandByName(String name) {
        List<Command> commands = new ArrayList<Command>();
        for (CommandData commandData : commandDataList) {
            if( name != null && name.equals(commandData.getCommand().getName()) ){
                commands.add(commandData.getCommand());
            }
        }
        return commands;
    }

    public void reset() {
        for (CommandData commandData : commandDataList) {
            commandData.reset();
        }
        transformer.resetTargetCellRefs();
    }

}
