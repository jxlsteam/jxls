package com.jxls.writer.command;

import com.jxls.writer.Cell;
import com.jxls.writer.CellRange;
import com.jxls.writer.Pos;
import com.jxls.writer.Size;
import com.jxls.writer.transform.Transformer;

import java.util.ArrayList;
import java.util.List;

/**
 * @author Leonid Vysochyn
 * Date: 1/16/12 6:44 PM
 */
public class BaseArea implements Area {
    public static final BaseArea EMPTY_AREA = new BaseArea(new Cell(0,0), Size.ZERO_SIZE);

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
        cellRange = new CellRange(startCell, initialSize.getWidth(), initialSize.getHeight());
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

    public void setTransformer(Transformer transformer) {
        this.transformer = transformer;
    }

    public Size applyAt(Cell cell, Context context) {
        int widthDelta = 0;
        int heightDelta = 0;
        for (int i = 0; i < commandDataList.size(); i++) {
            CommandData commandData = commandDataList.get(i);
            Cell newCell = new Cell(commandData.getStartPos().getCol() + cell.getCol(), commandData.getStartPos().getRow() + cell.getRow(), cell.getSheetIndex());
            Size initialSize = commandData.getCommand().getInitialSize();
            cellRange.excludeCells(commandData.getStartPos().getCol(), commandData.getStartPos().getCol() + commandData.getCommand().getInitialSize().getWidth() - 1,
                    commandData.getStartPos().getRow(), commandData.getStartPos().getRow() + commandData.getCommand().getInitialSize().getHeight() - 1);
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
                        data.setStartPos(new Pos(data.getStartPos().getCol(), data.getStartPos().getRow() + heightChange));
                    }else
                    if( newCol > newCell.getCol() && ( (newRow >= newCell.getRow() && newRow <= newCell.getRow() + newSize.getHeight()) ||
                   ( newRow + command.getInitialSize().getHeight() >= newCell.getRow() && newRow + command.getInitialSize().getHeight() <= newCell.getRow() + newSize.getHeight()) ||
                    newCell.getRow() >= newRow && newCell.getRow() <= newRow + command.getInitialSize().getHeight()) ){
                        data.setStartPos(new Pos(data.getStartPos().getCol() + widthChange, data.getStartPos().getRow()));
                    }
                }
            }
        }
        for(int x = 0; x < initialSize.getWidth(); x++){
            for(int y = 0; y < initialSize.getHeight(); y++){
                if( !cellRange.isExcluded(x,y) ){
                    Cell relativeCell = cellRange.getCell(x, y);
                    Cell srcCell = new Cell(startCell.getCol() + x, startCell.getRow() + y);
                    Cell targetCell = new Cell(relativeCell.getCol() + cell.getCol(), relativeCell.getRow() + cell.getRow());
                    transformer.transform(srcCell, targetCell, context);
                }
            }
        }
        return new Size(initialSize.getWidth() + widthDelta, initialSize.getHeight() + heightDelta);
    }

    public Size getSize(Context context) {
        Size newSize = new Size(initialSize.getWidth(), initialSize.getHeight());
        for(CommandData commandData: commandDataList){
            Size startSize = commandData.getCommand().getInitialSize();
            Size endSize = commandData.getCommand().getSize(context);
            if( startSize == null ){
                startSize = Size.ZERO_SIZE;
            }
            if( endSize == null ){
                endSize = Size.ZERO_SIZE;
            }
            int widthChange = endSize.getWidth() - startSize.getWidth();
            int heightChange = endSize.getHeight() - startSize.getHeight();
            newSize.append(widthChange, heightChange);
        }
        return newSize;
    }

    public Cell getStartCell() {
        return startCell;
    }

    public Size getInitialSize() {
        return initialSize;
    }
    
}
