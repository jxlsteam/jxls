package com.jxls.writer.command;

import com.jxls.writer.Pos;
import com.jxls.writer.Size;
import com.jxls.writer.transform.Transformer;

import java.util.ArrayList;
import java.util.List;

/**
 * @author Leonid Vysochyn
 * Date: 1/16/12 6:44 PM
 */
public class BaseCommand implements Command {
    public static final BaseCommand EMPTY_COMMAND = new BaseCommand(new Pos(0,0), Size.ZERO_SIZE);

    List<Command> commands;
    Transformer transformer;
    
    Pos pos;
    Size initialSize;

    public BaseCommand(Pos pos, Size initialSize, List<Command> commands, Transformer transformer) {
        this.pos = pos;
        this.initialSize = initialSize;
        this.commands = commands != null ? commands : new ArrayList<Command>();
        this.transformer = transformer;
    }

    public BaseCommand(Pos pos, Size initialSize) {
        this(pos, initialSize, null, null);
    }

    public BaseCommand(Pos pos, Size initialSize, Transformer transformer) {
        this(pos, initialSize, null, transformer);
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
    public Size applyAt(Pos pos, Context context) {
        int xDelta = pos.getX() - this.pos.getX();
        int yDelta = pos.getY() - this.pos.getY();
        for(Command command: commands){
            Pos newPos = new Pos(command.getPos().getX() + xDelta, command.getPos().getY() + yDelta);
            command.applyAt(newPos, context);
        }
        for(int x = this.pos.getX(), maxX = this.pos.getX() + initialSize.getWidth(); x < maxX; x++){
            for(int y = this.pos.getY(), maxY = this.pos.getY() + initialSize.getHeight(); y < maxY; y++){
                Pos origPos = new Pos(x,y);
                Pos newPos = new Pos(x + xDelta, y + yDelta);
                transformer.transform(origPos, newPos, context);
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
    public Pos getPos() {
        return pos;
    }

    @Override
    public Size getInitialSize() {
        return initialSize;
    }
    
}
