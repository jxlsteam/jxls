package com.jxls.writer.command;

import com.jxls.writer.Cell;
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
            //Command c = commandData.getCommand();
            Size newSize = commandData.getCommand().applyAt(newCell, context);
            int widthChange = newSize.getWidth() - initialSize.getWidth();
            int heightChange = newSize.getHeight() - initialSize.getHeight();
            if( widthChange != 0 || heightChange != 0){
                widthDelta += widthChange;
                heightDelta += heightChange;
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
        for(int x = startCell.getCol(), maxX = startCell.getCol() + initialSize.getWidth(); x < maxX; x++){
            for(int y = startCell.getRow(), maxY = startCell.getRow() + initialSize.getHeight(); y < maxY; y++){
                Cell origCell = new Cell(x,y, startCell.getSheetIndex());
                Cell newCell = new Cell(x - startCell.getCol() + cell.getCol(), y - startCell.getRow() + cell.getRow(), cell.getSheetIndex());
                transformer.transform(origCell, newCell, context);
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
