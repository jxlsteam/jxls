package com.jxls.writer.command;

import com.jxls.writer.Cell;
import com.jxls.writer.Size;
import com.jxls.writer.transform.Transformer;

import java.util.ArrayList;
import java.util.List;

/**
 * @author Leonid Vysochyn
 * Date: 1/16/12 6:44 PM
 */
public class BaseCommand implements Command {
    public static final BaseCommand EMPTY_COMMAND = new BaseCommand(new Cell(0,0), Size.ZERO_SIZE);

    List<Command> commands;
    Transformer transformer;
    
    Cell cell;
    Size initialSize;

    public BaseCommand(Cell cell, Size initialSize, List<Command> commands, Transformer transformer) {
        this.cell = cell;
        this.initialSize = initialSize;
        this.commands = commands != null ? commands : new ArrayList<Command>();
        this.transformer = transformer;
    }

    public BaseCommand(Cell cell, Size initialSize) {
        this(cell, initialSize, null, null);
    }

    public BaseCommand(Cell cell, Size initialSize, Transformer transformer) {
        this(cell, initialSize, null, transformer);
    }

    public void addCommand(Command command){
        commands.add(command);
    }

    public Transformer getTransformer() {
        return transformer;
    }

    public void setTransformer(Transformer transformer) {
        this.transformer = transformer;
    }

    @Override
    public Size applyAt(Cell cell, Context context) {
        int colDelta = cell.getCol() - this.cell.getCol();
        int rowDelta = cell.getRow() - this.cell.getRow();
        for(Command command: commands){
            Cell newCell = new Cell(command.getStartCell().getCol() + colDelta, command.getStartCell().getRow() + rowDelta);
            command.applyAt(newCell, context);
        }
        for(int x = this.cell.getCol(), maxX = this.cell.getCol() + initialSize.getWidth(); x < maxX; x++){
            for(int y = this.cell.getRow(), maxY = this.cell.getRow() + initialSize.getHeight(); y < maxY; y++){
                Cell origCell = new Cell(x,y);
                Cell newCell = new Cell(x + colDelta, y + rowDelta);
                transformer.transform(origCell, newCell, context);
            }
        }
        return null;
    }

    @Override
    public Size getSize(Context context) {
        Size newSize = new Size(initialSize.getWidth(), initialSize.getHeight());
        for(Command command: commands){
            Size startSize = command.getInitialSize();
            Size endSize = command.getSize(context);
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

    @Override
    public Cell getStartCell() {
        return cell;
    }

    @Override
    public Size getInitialSize() {
        return initialSize;
    }
    
}
