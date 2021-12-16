package org.jxls.area;

import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

import org.jxls.command.Command;
import org.jxls.common.AreaListener;
import org.jxls.common.AreaRef;
import org.jxls.common.CellData;
import org.jxls.common.CellRange;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;
import org.jxls.common.cellshift.AdjacentCellShiftStrategy;
import org.jxls.common.cellshift.CellShiftStrategy;
import org.jxls.common.cellshift.InnerCellShiftStrategy;
import org.jxls.formula.FormulaProcessor;
import org.jxls.formula.StandardFormulaProcessor;
import org.jxls.transform.Transformer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Core implementation of {@link Area} interface
 *
 * @author Leonid Vysochyn
 */
public class XlsArea implements Area {
    public static final XlsArea EMPTY_AREA = new XlsArea(new CellRef(null, 0, 0), Size.ZERO_SIZE);
    private static Logger logger = LoggerFactory.getLogger(XlsArea.class);

    private List<CommandData> commandDataList = new ArrayList<CommandData>();
    private Transformer transformer;
    private Command parentCommand;
    private CellRange cellRange;
    private CellRef startCellRef;
    private Size size;
    private List<AreaListener> areaListeners = new ArrayList<AreaListener>();
    private boolean cellsCleared = false;
    private FormulaProcessor formulaProcessor = new StandardFormulaProcessor();
    // default cell shift strategy
    private CellShiftStrategy cellShiftStrategy = new InnerCellShiftStrategy();
    private final CellShiftStrategy innerCellShiftStrategy = new InnerCellShiftStrategy();
    private final CellShiftStrategy adjacentCellShiftStrategy = new AdjacentCellShiftStrategy();

    public XlsArea(AreaRef areaRef, Transformer transformer) {
        CellRef startCell = areaRef.getFirstCellRef();
        CellRef endCell = areaRef.getLastCellRef();
        this.startCellRef = startCell;
        this.size = new Size(endCell.getCol() - startCell.getCol() + 1, endCell.getRow() - startCell.getRow() + 1);
        this.transformer = transformer;
    }

    public XlsArea(String areaRef, Transformer transformer) {
        this(new AreaRef(areaRef), transformer);
    }

    public XlsArea(CellRef startCell, CellRef endCell, Transformer transformer) {
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
    public Command getParentCommand() {
        return parentCommand;
    }

    @Override
    public void setParentCommand(Command command) {
        this.parentCommand = command;
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

    @Override
    public void addCommand(AreaRef areaRef, Command command) {
        AreaRef thisAreaRef = new AreaRef(startCellRef, size);
        if (!thisAreaRef.contains(areaRef)) {
            throw new IllegalArgumentException("Cannot add command '" + command.getName() + "' to area " + thisAreaRef + " at " + areaRef);
        }
        commandDataList.add(new CommandData(areaRef, command));
    }

    public void addCommand(String areaRef, Command command) {
        commandDataList.add(new CommandData(areaRef, command));
    }

    @Override
    public List<CommandData> getCommandDataList() {
        return commandDataList;
    }

    @Override
    public Transformer getTransformer() {
        return transformer;
    }

    public void setTransformer(Transformer transformer) {
        this.transformer = transformer;
    }

    private void excludeCells(CellRef startCellRef, Size size) {
        cellRange.excludeCells(
                startCellRef.getCol() - this.startCellRef.getCol(),
                startCellRef.getCol() - this.startCellRef.getCol() + size.getWidth() - 1,
                startCellRef.getRow() - this.startCellRef.getRow(),
                startCellRef.getRow() - this.startCellRef.getRow() + size.getHeight() - 1
        );
    }

    private void createCellRange() {
        cellRange = new CellRange(startCellRef, size.getWidth(), size.getHeight());
        for (CommandData commandData : commandDataList) {
            CellRef startCellRef = commandData.getSourceStartCellRef();
            Size size = commandData.getSourceSize();
            if (commandData.getCommand().getLockRange()) {
                excludeCells(startCellRef, size);
            }
        }
    }

    // TODO Method too long
    @Override
    public Size applyAt(CellRef cellRef, Context context) {
        if (this == XlsArea.EMPTY_AREA) {
            return Size.ZERO_SIZE;
        }
        logger.debug("Applying XlsArea at {}", cellRef);
        fireBeforeApplyEvent(cellRef, context);
        createCellRange();
        AreaRef commandsArea = transformTopStaticArea(cellRef, context);
        for (int i = 0; i < commandDataList.size(); i++) {
            cellRange.resetChangeMatrix();
            CommandData commandData = commandDataList.get(i);
            String shiftMode = commandData.getCommand().getShiftMode();
            CellShiftStrategy commandCellShiftStrategy = detectCellShiftStrategy(shiftMode);
            cellRange.setCellShiftStrategy(commandCellShiftStrategy);
            CellRef commandStartCellRef = commandData.getStartCellRef();
            Size commandInitialSize = commandData.getSize();
            int startCol = commandStartCellRef.getCol() - startCellRef.getCol();
            int startRow = commandStartCellRef.getRow() - startCellRef.getRow();
            CellRef newCell = new CellRef(cellRef.getSheetName(), startRow + cellRef.getRow(), startCol + cellRef.getCol());
            Size commandNewSize = commandData.getCommand().applyAt(newCell, context);
            int widthChange = commandNewSize.getWidth() - commandInitialSize.getWidth();
            int heightChange = commandNewSize.getHeight() - commandInitialSize.getHeight();
            int endCol = startCol + commandInitialSize.getWidth() - 1;
            int endRow = startRow + commandInitialSize.getHeight() - 1;
            if (heightChange != 0) {
                cellRange.shiftCellsWithColBlock(startCol, endCol, endRow, heightChange, true);
                Set<CommandData> commandsToShift = findCommandsForVerticalShift(commandDataList.subList(i+1, commandDataList.size()),
                        startCol, endCol, endRow, heightChange);
                for (CommandData commandDataToShift : commandsToShift) {
                    CellRef commandDataStartCellRef = commandDataToShift.getStartCellRef();
                    int relativeRow = commandDataStartCellRef.getRow() - startCellRef.getRow();
                    int relativeStartCol = commandDataStartCellRef.getCol() - startCellRef.getCol();
                    int relativeEndCol = relativeStartCol + commandDataToShift.getSize().getWidth() - 1;
                    cellRange.shiftCellsWithColBlock(relativeStartCol, relativeEndCol,
                            relativeRow + commandDataToShift.getSize().getHeight() - 1, heightChange, false);
                    commandDataToShift.setStartCellRef(
                            new CellRef(commandStartCellRef.getSheetName(),
                                    commandDataStartCellRef.getRow() + heightChange,
                                    commandDataStartCellRef.getCol()));
                    if (heightChange < 0) {
                        CellRef initialStartCellRef = commandDataToShift.getSourceStartCellRef();
                        Size initialSize = commandDataToShift.getSourceSize();
                        int initialStartRow = initialStartCellRef.getRow() - startCellRef.getRow();
                        int initialEndRow = initialStartRow + initialSize.getHeight() - 1;
                        int initialStartCol = initialStartCellRef.getCol() - startCellRef.getCol();
                        int initialEndCol = initialStartCol + initialSize.getWidth() - 1;
                        cellRange.clearCells(initialStartCol, initialEndCol, initialStartRow, initialEndRow);
                    }
                }
            }
            if (widthChange != 0) {
                cellRange.shiftCellsWithRowBlock(startRow,
                        endRow,
                        endCol, widthChange, true);
                Set<CommandData> commandsToShift = findCommandsForHorizontalShift(
                        commandDataList.subList(i + 1, commandDataList.size()), startRow, endRow, endCol, widthChange);
                for (CommandData commandDataToShift : commandsToShift) {
                    CellRef commandDataStartCellRef = commandDataToShift.getStartCellRef();
                    int relativeCol = commandDataStartCellRef.getCol() - startCellRef.getCol();
                    int relativeStartRow = commandDataStartCellRef.getRow() - startCellRef.getRow();
                    int relativeEndRow = relativeStartRow + commandDataToShift.getSize().getHeight() - 1;
                    cellRange.shiftCellsWithRowBlock(relativeStartRow, relativeEndRow,
                            relativeCol + commandDataToShift.getSize().getWidth() - 1, widthChange, false);
                    commandDataToShift.setStartCellRef(
                            new CellRef(commandStartCellRef.getSheetName(),
                                    commandDataStartCellRef.getRow(),
                                    commandDataStartCellRef.getCol() + widthChange));
                    if (widthChange < 0) {
                        CellRef initialStartCellRef = commandDataToShift.getSourceStartCellRef();
                        Size initialSize = commandDataToShift.getSourceSize();
                        int initialStartRow = initialStartCellRef.getRow() - startCellRef.getRow();
                        int initialEndRow = initialStartRow + initialSize.getHeight() - 1;
                        int initialStartCol = initialStartCellRef.getCol() - startCellRef.getCol();
                        int initialEndCol = initialStartCellRef.getCol() + initialSize.getWidth() - 1;
                        cellRange.clearCells(initialStartCol, initialEndCol, initialStartRow, initialEndRow);
                    }
                }
            }
        }
        transformStaticCells(cellRef, context, commandsArea);
        fireAfterApplyEvent(cellRef, context);
        Size finalSize = new Size(cellRange.calculateWidth(), cellRange.calculateHeight());
        AreaRef newAreaRef = new AreaRef(cellRef, finalSize);
        updateCellDataFinalAreaForFormulaCells(newAreaRef);
        for (CommandData commandData : commandDataList) {
            commandData.resetStartCellAndSize();
        }
        return finalSize;
    }

    private void transformStaticCells(CellRef cellRef, Context context, AreaRef commandsArea) {
        int relativeStartRow = commandsArea.getFirstCellRef().getRow();
        int relativeStartCol = commandsArea.getFirstCellRef().getCol() + 1;
        if (transformer.isForwardOnly()) {
            relativeStartRow = commandsArea.getLastCellRef().getRow() + 1;
            relativeStartCol = 0;
        }
        transformStaticCells(cellRef, context, relativeStartRow, relativeStartCol);
    }

    // TODO similar code to findCommandsForVerticalShift() ?
    private Set<CommandData> findCommandsForHorizontalShift(List<CommandData> commandList, int startRow, int endRow, int shiftingCol, int widthChange) {
        Set<CommandData> result = new LinkedHashSet<>(commandList.size());
        for (int i = 0, commandListSize = commandList.size(); i < commandListSize; i++) {
            CommandData commandData = commandList.get(i);
            if (result.contains(commandData)) continue; // Should be safe with default equals & hashcode (as references are not changing)

            CellRef commandDataStartCellRef = commandData.getStartCellRef();
            int relativeCol = commandDataStartCellRef.getCol() - startCellRef.getCol();
            int relativeStartRow = commandDataStartCellRef.getRow() - startCellRef.getRow();
            int relativeEndRow = relativeStartRow + commandData.getSize().getHeight() - 1;
            if (relativeCol > shiftingCol) {
                boolean isShiftingNeeded = false;
                if (widthChange > 0) {
                    if ((relativeStartRow >= startRow && relativeStartRow <= endRow)
                            || (relativeEndRow >= startRow && relativeEndRow <= endRow)
                            || (startRow >= relativeStartRow && startRow <= relativeEndRow)) {
                        isShiftingNeeded = true;
                    }
                } else {
                    if (relativeStartRow >= startRow && relativeEndRow <= endRow
                            && isNoHighCommandsInArea(commandList, shiftingCol + 1, relativeCol - 1, startRow, endRow)) {
                        isShiftingNeeded = true;
                    }
                }
                if (isShiftingNeeded) {
                    result.add(commandData);
                    Set<CommandData> dependentCommands = findCommandsForHorizontalShift(
                            commandList.subList(i + 1, commandList.size()),
                            relativeStartRow,
                            relativeEndRow,
                            relativeCol + commandData.getSize().getWidth() - 1,
                            widthChange);
                    result.addAll(dependentCommands);
                }
            }
        }
        return result;
    }

    // TODO similar code to isNoWideCommandsInArea() ?  (Other duplicate code in this file?)
    private boolean isNoHighCommandsInArea(List<CommandData> commandList, int startCol, int endCol, int startRow, int endRow) {
        for (CommandData commandData : commandList) {
            CellRef commandDataStartCellRef = commandData.getStartCellRef();
            int relativeCol = commandDataStartCellRef.getCol() - startCellRef.getCol();
            int relativeEndCol = relativeCol + commandData.getSize().getWidth() - 1;
            int relativeStartRow = commandDataStartCellRef.getRow() - startCellRef.getRow();
            int relativeEndRow = relativeStartRow + commandData.getSize().getHeight() - 1;

            if (relativeCol >= startCol && relativeEndCol <= endCol
                    && ((relativeStartRow < startRow && relativeEndRow >= startRow) || (relativeEndRow > endRow && relativeStartRow <= endRow))) {
                return false;
            }
        }
        return true;
    }

    private Set<CommandData> findCommandsForVerticalShift(List<CommandData> commandList, int startCol, int endCol, int shiftingRow, int heightChange) {
        Set<CommandData> result = new LinkedHashSet<>(commandList.size());
        int commandListSize = commandList.size();
        for (int i = 0; i < commandListSize; i++) {
            CommandData commandData = commandList.get(i);
            if (result.contains(commandData)) continue; // Should be safe with default equals & hashcode (as references are not changing)

            CellRef commandDataStartCellRef = commandData.getStartCellRef();
            int relativeRow = commandDataStartCellRef.getRow() - startCellRef.getRow();
            int relativeStartCol = commandDataStartCellRef.getCol() - startCellRef.getCol();
            int relativeEndCol = relativeStartCol + commandData.getSize().getWidth() - 1;
            if (relativeRow > shiftingRow) {
                boolean isShiftingNeeded = false;
                if (heightChange > 0) {
                    if ((relativeStartCol >= startCol && relativeStartCol <= endCol)
                            || (relativeEndCol >= startCol && relativeEndCol <= endCol)
                            || (startCol >= relativeStartCol && startCol <= relativeEndCol)) {
                        isShiftingNeeded = true;
                    }
                } else {
                    if (relativeStartCol >= startCol && relativeEndCol <= endCol && isNoWideCommandsInArea(commandList, startCol, endCol, shiftingRow + 1, relativeRow - 1)) {
                        isShiftingNeeded = true;
                    }
                }
                if (isShiftingNeeded) {
                    result.add(commandData);
                    Set<CommandData> dependentCommands = findCommandsForVerticalShift(
                            commandList.subList(i + 1, commandListSize),
                            relativeStartCol,
                            relativeEndCol,
                            relativeRow + commandData.getSize().getHeight() - 1,
                            heightChange);
                    result.addAll(dependentCommands);
                }
            }
        }
        return result;
    }

    private boolean isNoWideCommandsInArea(List<CommandData> commandList, int startCol, int endCol, int startRow, int endRow) {
        for (CommandData commandData : commandList) {
            CellRef commandDataStartCellRef = commandData.getStartCellRef();
            int relativeRow = commandDataStartCellRef.getRow() - startCellRef.getRow();
            int relativeEndRow = relativeRow + commandData.getSize().getHeight() - 1;
            int relativeStartCol = commandDataStartCellRef.getCol() - startCellRef.getCol();
            int relativeEndCol = relativeStartCol + commandData.getSize().getWidth() - 1;
            if( relativeRow >= startRow && relativeEndRow <= endRow
                    && ((relativeStartCol < startCol && relativeEndCol >= startCol) || (relativeEndCol > endCol && relativeStartCol <= endCol))) {
                return false;
            }
        }
        return true;
    }

    private CellShiftStrategy detectCellShiftStrategy(String shiftMode) {
        if (shiftMode != null && Command.ADJACENT_SHIFT_MODE.equalsIgnoreCase(shiftMode)) {
            return adjacentCellShiftStrategy;
        } else {
            return innerCellShiftStrategy;
        }
    }

    private void updateCellDataFinalAreaForFormulaCells(AreaRef newAreaRef) {
        String sheetName = startCellRef.getSheetName();
        int offsetRow = startCellRef.getRow();
        int startCol = startCellRef.getCol();
        for (int col = 0; col < size.getWidth(); col++) {
            for (int row = 0; row < size.getHeight(); row++) {
                if (!cellRange.isExcluded(row, col)) {
                    CellRef srcCell = new CellRef(sheetName, offsetRow + row, startCol + col);
                    CellData cellData = transformer.getCellData(srcCell);
                    if (cellData != null && cellData.isFormulaCell()) {
                        cellData.addTargetParentAreaRef(newAreaRef);
                    }
                }
            }
        }
    }

    private AreaRef transformTopStaticArea(CellRef cellRef, Context context) {
        CellRef topLeftCommandCell = findRelativeTopCommandCellRef();
        CellRef bottomRightCommandCell = findRelativeBottomCommandCellRef();
        int topStaticAreaLastRow = topLeftCommandCell.getRow() - 1;
        for (int col = 0; col < size.getWidth(); col++) {
            for (int row = 0; row <= topStaticAreaLastRow; row++) {
                transformStaticCell(cellRef, context, row, col);
            }
        }
        // update static cells before the command cell for the first top left command
        for (int col = 0; col < topLeftCommandCell.getCol(); col++) {
            if (cellRange.contains(topLeftCommandCell.getRow(), col)) {
                transformStaticCell(cellRef, context, topLeftCommandCell.getRow(), col);
            }
        }
        if (parentCommand == null) {
            updateRowHeights(cellRef, 0, topStaticAreaLastRow);
        }
        return new AreaRef(topLeftCommandCell, bottomRightCommandCell);
    }

    private void transformStaticCell(CellRef cellRef, Context context, int row, int col) {
        String sheetName = startCellRef.getSheetName();
        int startRow = startCellRef.getRow();
        int startCol = startCellRef.getCol();
        if (!cellRange.isExcluded(row, col)) {
            CellRef relativeCell = cellRange.getCell(row, col);
            CellRef srcCell = new CellRef(sheetName, startRow + row, startCol + col);
            CellRef targetCell = new CellRef(cellRef.getSheetName(), relativeCell.getRow() + cellRef.getRow(), relativeCell.getCol() + cellRef.getCol());
            fireBeforeTransformCell(srcCell, targetCell, context);
            try {
                updateCellDataArea(srcCell, targetCell, context);
                boolean updateRowHeight = parentCommand != null;
                transformer.transform(srcCell, targetCell, context, updateRowHeight);
            } catch (Exception e) {
                logger.error("Failed to transform " + srcCell + " into " + targetCell, e);
            }
            fireAfterTransformCell(srcCell, targetCell, context);
        }
    }

    private void updateRowHeights(CellRef areaStartCellRef, int relativeStartRow, int relativeEndRow) {
        if (transformer == null) return;
        for (int relativeSrcRow = relativeStartRow; relativeSrcRow <= relativeEndRow; relativeSrcRow++) {
            if (!cellRange.containsCommandsInRow(relativeSrcRow)) {
                int relativeTargetRow = cellRange.findTargetRow(relativeSrcRow);
                int targetRow = areaStartCellRef.getRow() + relativeTargetRow;
                int srcRow = startCellRef.getRow() + relativeSrcRow;
                try {
                    transformer.updateRowHeight(startCellRef.getSheetName(), srcRow, areaStartCellRef.getSheetName(), targetRow);
                } catch (Exception e) {
                    logger.error("Failed to update row height for src row={} and target row={}", relativeSrcRow, targetRow, e);
                }
            }
        }
    }

    private CellRef findRelativeTopCommandCellRef() {
        int topCommandRow = startCellRef.getRow() + size.getHeight();
        int topCommandCol = startCellRef.getCol() + size.getWidth();
        for (CommandData data : commandDataList) {
            if (data.getStartCellRef().getRow() <= topCommandRow && data.getStartCellRef().getCol() <= topCommandCol) {
                topCommandRow = data.getStartCellRef().getRow();
                topCommandCol = data.getStartCellRef().getCol();
            }
        }
        return new CellRef(topCommandRow - startCellRef.getRow(), topCommandCol - startCellRef.getCol());
    }

    private CellRef findRelativeBottomCommandCellRef() {
        int bottomCommandRow = startCellRef.getRow();
        int bottomCommandCol = startCellRef.getCol();
        for (CommandData data : commandDataList) {
            if (data.getStartCellRef().getRow() + data.getSize().getHeight() >= bottomCommandRow) {
                bottomCommandRow = data.getStartCellRef().getRow() + data.getSize().getHeight() - 1;
                bottomCommandCol = data.getStartCellRef().getCol() + data.getSize().getWidth() - 1;
            }
        }
        return new CellRef(bottomCommandRow - startCellRef.getRow(), bottomCommandCol - startCellRef.getCol());
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
        if (cellsCleared) return;
        String sheetName = startCellRef.getSheetName();
        int startRow = startCellRef.getRow();
        int startCol = startCellRef.getCol();
        for (int row = 0; row < size.getHeight(); row++) {
            for (int col = 0; col < size.getWidth(); col++) {
                CellRef cellRef = new CellRef(sheetName, startRow + row, startCol + col);
                transformer.clearCell(cellRef);
            }
        }
        transformer.resetArea(getAreaRef());
        cellsCleared = true;
    }

    private void transformStaticCells(CellRef cellRef, Context context, int relativeStartRow, int relativeStartCol) {
        String sheetName = startCellRef.getSheetName();
        int offsetRow = startCellRef.getRow();
        int startCol = startCellRef.getCol();
        int width = size.getWidth();
        int height = size.getHeight();
        for (int col = 0; col < width; col++) {
            for (int row = relativeStartRow; row < height; row++) {
                if (row == relativeStartRow && col < relativeStartCol) continue;
                if (!cellRange.isExcluded(row, col)) {
                    CellRef relativeCell = cellRange.getCell(row, col);
                    CellRef srcCell = new CellRef(sheetName, offsetRow + row, startCol + col);
                    CellRef targetCell = new CellRef(cellRef.getSheetName(), relativeCell.getRow() + cellRef.getRow(), relativeCell.getCol() + cellRef.getCol());
                    fireBeforeTransformCell(srcCell, targetCell, context);
                    try {
                        updateCellDataArea(srcCell, targetCell, context);
                        boolean updateRowHeight = parentCommand != null;
                        transformer.transform(srcCell, targetCell, context, updateRowHeight);
                    } catch (Exception e) {
                        logger.error("Failed to transform " + srcCell + " into " + targetCell, e);
                    }
                    fireAfterTransformCell(srcCell, targetCell, context);
                }
            }
        }
        if (parentCommand == null) {
            updateRowHeights(cellRef, relativeStartRow, height - 1);
        }
    }

    private void updateCellDataArea(CellRef srcCell, CellRef targetCell, Context context) {
        Context.Config config = context.getConfig();
        if (!config.isFormulaProcessingRequired()) return;
        CellData cellData = transformer.getCellData(srcCell);
        if (cellData != null) {
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

    @Override
    public CellRef getStartCellRef() {
        return startCellRef;
    }

    @Override
    public Size getSize() {
        return size;
    }

    @Override
    public AreaRef getAreaRef() {
        return new AreaRef(startCellRef, size);
    }

    @Override
    public void processFormulas() {
        formulaProcessor.processAreaFormulas(transformer, this);
    }

    @Override
    public void addAreaListener(AreaListener listener) {
        areaListeners.add(listener);
    }

    @Override
    public List<AreaListener> getAreaListeners() {
        return areaListeners;
    }

    @Override
    public List<Command> findCommandByName(String name) {
        List<Command> commands = new ArrayList<Command>();
        for (CommandData commandData : commandDataList) {
            if (name != null && name.equals(commandData.getCommand().getName())) {
                commands.add(commandData.getCommand());
            }
        }
        return commands;
    }

    @Override
    public void reset() {
        for (CommandData commandData : commandDataList) {
            commandData.reset();
        }
        transformer.resetTargetCellRefs();
    }
}
