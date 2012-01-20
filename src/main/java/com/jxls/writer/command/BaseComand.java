package com.jxls.writer.command;

import com.jxls.writer.Pos;
import com.jxls.writer.Size;

import java.util.ArrayList;
import java.util.List;

/**
 * @author Leonid Vysochyn
 * Date: 1/16/12 6:44 PM
 */
public class BaseComand implements Command {
    public static final BaseComand EMPTY_COMAND = new BaseComand(new Pos(0,0), Size.ZERO_SIZE);

    List<Command> commands;
    
    Pos pos;
    Size initialSize;

    public BaseComand(Pos pos, Size initialSize, List<Command> commands) {
        this.pos = pos;
        this.initialSize = initialSize;
        this.commands = commands != null ? commands : new ArrayList<Command>();
    }

    public BaseComand(Pos pos, Size initialSize) {
        this(pos, initialSize, null);
    }

    @Override
    public Size applyAt(Pos pos, Context context) {
        return null;
    }

    @Override
    public Size getSize(Context context) {
        Size newSize = new Size(initialSize.getWidth(), initialSize.getHeight());
        for(Command command: commands){
            Size startSize = command.getInitialSize();
            Size endSize = command.getSize(context);
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
