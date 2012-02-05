package com.jxls.writer.command;

import com.jxls.writer.*;
import com.jxls.writer.transform.Transformer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Set;

/**
 * @author Leonid Vysochyn
 * Date: 1/16/12 6:44 PM
 */
public class BaseArea implements Area {
    static Logger logger = LoggerFactory.getLogger(BaseArea.class);

    public static final BaseArea EMPTY_AREA = new BaseArea(new Pos(0, 0), Size.ZERO_SIZE);

    List<CommandData> commandDataList;
    Transformer transformer;
    
    CellRange cellRange;
    
    Pos startPos;
    Size initialSize;

    public BaseArea(Pos startPos, Size initialSize, List<CommandData> commandDataList, Transformer transformer) {
        this.startPos = startPos;
        this.initialSize = initialSize;
        this.commandDataList = commandDataList != null ? commandDataList : new ArrayList<CommandData>();
        this.transformer = transformer;
    }

    public BaseArea(Pos startPos, Size initialSize) {
        this(startPos, initialSize, null, null);
    }

    public BaseArea(Pos startPos, Size initialSize, Transformer transformer) {
        this(startPos, initialSize, null, transformer);
    }

    public void addCommand(Pos pos, Command command) {
        commandDataList.add(new CommandData(pos, command));
    }

    public Transformer getTransformer() {
        return transformer;
    }

    public void processFormulas() {
        Set<CellData> formulaCells = transformer.getFormulaCells();
        for (CellData formulaCellData : formulaCells) {
            String targetFormulaString = formulaCellData.getFormula();
            List<String> formulaCellRefs = Util.getFormulaCellRefs(formulaCellData.getFormula());
            for (String cellRef : formulaCellRefs) {
                Pos pos = new Pos(cellRef);
                Set<Pos> targetCellDataList = transformer.getTargetPos(pos.getSheet(), pos.getRow(), pos.getCol());
                String targetCellRef = Util.createTargetCellRef(targetCellDataList);
                targetFormulaString = targetFormulaString.replaceAll(cellRef, targetCellRef);
            }
            Set<Pos> targetFormulaCells = transformer.getTargetPos(formulaCellData.getSheet(), formulaCellData.getRow(), formulaCellData.getCol());
            for (Iterator<Pos> iterator = targetFormulaCells.iterator(); iterator.hasNext(); ) {
                Pos targetFormulaCell = iterator.next();
                transformer.updateFormulaCell(new Pos(targetFormulaCell.getSheet(), targetFormulaCell.getRow(), targetFormulaCell.getCol()), targetFormulaString);
            }

        }
    }


    public void setTransformer(Transformer transformer) {
        this.transformer = transformer;
    }
    
    private void createCellRange(){
        cellRange = new CellRange(startPos, initialSize.getWidth(), initialSize.getHeight());
        for(CommandData commandData: commandDataList){
            cellRange.excludeCells(commandData.getStartPos().getCol(), commandData.getStartPos().getCol() + commandData.getCommand().getInitialSize().getWidth()-1,
                    commandData.getStartPos().getRow(), commandData.getStartPos().getRow() + commandData.getCommand().getInitialSize().getHeight()-1);
        }
    }

    public Size applyAt(Pos pos, Context context) {
        logger.debug("Applying BaseArea at {} with {}", pos, context);
        int widthDelta = 0;
        int heightDelta = 0;
        createCellRange();
        for (int i = 0; i < commandDataList.size(); i++) {
            cellRange.resetChangeMatrix();
            CommandData commandData = commandDataList.get(i);
            Pos newCell = new Pos(pos.getSheet(), commandData.getStartPos().getRow() + pos.getRow(), commandData.getStartPos().getCol() + pos.getCol());
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
                    int newRow = data.getStartPos().getRow() + pos.getRow();
                    int newCol = data.getStartPos().getCol() + pos.getCol();
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
        transformStaticCells(pos, context);
        return new Size(initialSize.getWidth() + widthDelta, initialSize.getHeight() + heightDelta);
    }

    private void transformStaticCells(Pos Pos, Context context) {
        for(int x = 0; x < initialSize.getWidth(); x++){
            for(int y = 0; y < initialSize.getHeight(); y++){
                if( !cellRange.isExcluded(y, x) ){
                    Pos relativeCell = cellRange.getCell(y, x);
                    Pos srcCell = new Pos(startPos.getSheet(), startPos.getRow() + y, startPos.getCol() + x);
                    Pos targetCell = new Pos(Pos.getSheet(), relativeCell.getRow() + Pos.getRow(), relativeCell.getCol() + Pos.getCol());
                    transformer.transform(srcCell, targetCell, context);
                }
            }
        }
    }

    public Pos getStartPos() {
        return startPos;
    }

    public Size getInitialSize() {
        return initialSize;
    }
    
}
