package com.jxls.writer.command;

import com.jxls.writer.*;
import com.jxls.writer.transform.Transformer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.List;

/**
 * @author Leonid Vysochyn
 * Date: 1/16/12 6:44 PM
 */
public class BaseArea implements Area {
    static Logger logger = LoggerFactory.getLogger(BaseArea.class);

    public static final BaseArea EMPTY_AREA = new BaseArea(new Cell(0, 0), Size.ZERO_SIZE);

    List<CommandData> commandDataList;
    Transformer transformer;
    
    CellRange cellRange;
    
    Cell startCell;
    Size initialSize;

    public BaseArea(Cell startCell, Size initialSize, List<CommandData> commandDataList, Transformer transformer) {
        this.startCell = startCell;
        this.initialSize = initialSize;
        this.commandDataList = commandDataList != null ? commandDataList : new ArrayList<CommandData>();
        this.transformer = transformer;
    }

    public BaseArea(Cell startCell, Size initialSize) {
        this(startCell, initialSize, null, null);
    }

    public BaseArea(Cell startCell, Size initialSize, Transformer transformer) {
        this(startCell, initialSize, null, transformer);
    }

    public void addCommand(Pos pos, Command command) {
        commandDataList.add(new CommandData(pos, command));
    }

    public Transformer getTransformer() {
        return transformer;
    }

    public void processFormulas() {
        List<CellData> formulaCells = transformer.getFormulaCells();
    }

    public void setTransformer(Transformer transformer) {
        this.transformer = transformer;
    }
    
    private void createCellRange(){
        cellRange = new CellRange(startCell, initialSize.getWidth(), initialSize.getHeight());
        for(CommandData commandData: commandDataList){
            cellRange.excludeCells(commandData.getStartPos().getCol(), commandData.getStartPos().getCol() + commandData.getCommand().getInitialSize().getWidth()-1,
                    commandData.getStartPos().getRow(), commandData.getStartPos().getRow() + commandData.getCommand().getInitialSize().getHeight()-1);
        }
    }

    public Size applyAt(Cell cell, Context context) {
        logger.debug("Applying BaseArea at {} with {}", cell, context);
        int widthDelta = 0;
        int heightDelta = 0;
        createCellRange();
        for (int i = 0; i < commandDataList.size(); i++) {
            cellRange.resetChangeMatrix();
            CommandData commandData = commandDataList.get(i);
            Cell newCell = new Cell(cell.getSheetIndex(), commandData.getStartPos().getRow() + cell.getRow(), commandData.getStartPos().getCol() + cell.getCol());
            Size initialSize = commandData.getCommand().getInitialSize();
            Size newSize = commandData.getCommand().applyAt(newCell, context);
            int widthChange = newSize.getWidth() - initialSize.getWidth();
            int heightChange = newSize.getHeight() - initialSize.getHeight();
            if( widthChange != 0 || heightChange != 0){
                widthDelta += widthChange;
                heightDelta += heightChange;
                if( widthChange != 0 ){
                    cellRange.shiftCellsWithRowBlock(commandData.getStartPos().getRow(),
                            commandData.getStartPos().getRow() + commandData.getCommand().getInitialSize().getHeight(),
                            commandData.getStartPos().getCol() + initialSize.getWidth(), widthChange);
                }
                if( heightChange != 0 ){
                    cellRange.shiftCellsWithColBlock(commandData.getStartPos().getCol(),
                            commandData.getStartPos().getCol() + newSize.getWidth()-1, commandData.getStartPos().getRow() + commandData.getCommand().getInitialSize().getHeight()-1, heightChange);
                }
                for (int j = i + 1; j < commandDataList.size(); j++) {
                    CommandData data = commandDataList.get(j);
                    Command command = data.getCommand();
                    int newRow = data.getStartPos().getRow() + cell.getRow();
                    int newCol = data.getStartPos().getCol() + cell.getCol();
                    if(newRow > newCell.getRow() && ((newCol >= newCell.getCol() && newCol <= newCell.getCol() + newSize.getWidth()) ||
                            (newCol + command.getInitialSize().getWidth() >= newCell.getCol() && newCol + command.getInitialSize().getWidth() <= newCell.getCol() + newSize.getWidth()) ||
                            (newCell.getCol() >= newCol && newCell.getCol() <= newCol + command.getInitialSize().getWidth() )
                    )){
                        cellRange.shiftCellsWithColBlock(data.getStartPos().getCol(),
                                data.getStartPos().getCol() + data.getCommand().getInitialSize().getWidth()-1, data.getStartPos().getRow() + data.getCommand().getInitialSize().getHeight()-1, heightChange);
                        data.setStartPos(new Pos(data.getStartPos().getRow() + heightChange, data.getStartPos().getCol()));
                    }else
                    if( newCol > newCell.getCol() && ( (newRow >= newCell.getRow() && newRow <= newCell.getRow() + newSize.getHeight()) ||
                   ( newRow + command.getInitialSize().getHeight() >= newCell.getRow() && newRow + command.getInitialSize().getHeight() <= newCell.getRow() + newSize.getHeight()) ||
                    newCell.getRow() >= newRow && newCell.getRow() <= newRow + command.getInitialSize().getHeight()) ){
                        cellRange.shiftCellsWithRowBlock(data.getStartPos().getRow(),
                                data.getStartPos().getRow() + data.getCommand().getInitialSize().getHeight()-1,
                                data.getStartPos().getCol() + initialSize.getWidth(), widthChange);
                        data.setStartPos(new Pos(data.getStartPos().getRow(), data.getStartPos().getCol() + widthChange));
                    }
                }
            }
        }
        transformStaticCells(cell, context);
        return new Size(initialSize.getWidth() + widthDelta, initialSize.getHeight() + heightDelta);
    }

    private void transformStaticCells(Cell cell, Context context) {
        for(int x = 0; x < initialSize.getWidth(); x++){
            for(int y = 0; y < initialSize.getHeight(); y++){
                if( !cellRange.isExcluded(y, x) ){
                    Cell relativeCell = cellRange.getCell(y, x);
                    Cell srcCell = new Cell(startCell.getSheetIndex(), startCell.getRow() + y, startCell.getCol() + x);
                    Cell targetCell = new Cell(cell.getSheetIndex(), relativeCell.getRow() + cell.getRow(), relativeCell.getCol() + cell.getCol());
                    transformer.transform(srcCell, targetCell, context);
                }
            }
        }
    }

    public Cell getStartCell() {
        return startCell;
    }

    public Size getInitialSize() {
        return initialSize;
    }
    
}
