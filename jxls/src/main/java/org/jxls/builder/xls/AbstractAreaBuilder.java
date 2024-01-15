package org.jxls.builder.xls;

import java.util.ArrayList;
import java.util.List;

import org.jxls.area.Area;
import org.jxls.area.CommandData;
import org.jxls.area.XlsArea;
import org.jxls.builder.AreaBuilder;
import org.jxls.common.AreaRef;
import org.jxls.common.CellData;
import org.jxls.transform.Transformer;

public abstract class AbstractAreaBuilder implements AreaBuilder {
    
    /**
     * Builds a list of {@link org.jxls.area.XlsArea} objects defined by top level AreaCommand markup ("jx:area")
     * containing a tree of all nested commands
     * 
     * @param transformer -
     * @param clearTemplateCells default is true
     * @return Area list
     */
    @Override
    public List<Area> build(Transformer transformer, boolean clearTemplateCells) {
        List<CommandData> allCommands = new ArrayList<>();
        List<Area> allAreas = new ArrayList<>();
        List<Area> userAreas = new ArrayList<>();
        transformer.getCommentedCells().forEach(d -> {
            List<CommandData> commandDatas = buildCommands(transformer, d, d.getCellComment());
            processCommandData(commandDatas, allCommands, allAreas, userAreas, transformer);
        });
        return processCommands(allCommands, allAreas, userAreas, clearTemplateCells);
    }

    protected abstract List<CommandData> buildCommands(Transformer transformer, CellData cellData, String text);
    
    protected void processCommandData(List<CommandData> commandDatas, List<CommandData> allCommands, List<Area> allAreas,
            List<Area> userAreas, Transformer transformer) {
        for (CommandData commandData : commandDatas) {
            if (commandData.getCommand() instanceof AreaCommand) {
                XlsArea userArea = new XlsArea(commandData.getAreaRef(), transformer);
                allAreas.add(userArea);
                userAreas.add(userArea);
            } else {
                List<Area> areas = commandData.getCommand().getAreaList();
                allAreas.addAll(areas);
                allCommands.add(commandData);
            }
        }
    }

    protected List<Area> processCommands(List<CommandData> allCommands, List<Area> allAreas, List<Area> userAreas, boolean clearTemplateCells) {
        for (int i = 0; i < allCommands.size(); i++) {
            CommandData commandData = allCommands.get(i);
            AreaRef commandAreaRef = commandData.getAreaRef();
            List<Area> commandAreas = commandData.getCommand().getAreaList();
            Area minArea = null;
            List<Area> minAreas = new ArrayList<>();
            for (Area area : allAreas) {
                if (commandAreas.contains(area) || !area.getAreaRef().contains(commandAreaRef)) continue;
                boolean belongsToNextCommand = false;
                for (int j = i + 1; j < allCommands.size(); j++) {
                    CommandData nextCommand = allCommands.get(j);
                    if (nextCommand.getCommand().getAreaList().contains(area)) {
                        belongsToNextCommand = true;
                        break;
                    }
                }
                if (belongsToNextCommand || (minArea != null && !minArea.getAreaRef().contains(area.getAreaRef()))) continue;
                if (minArea != null && minArea.equals(area)) {
                    minAreas.add(area);
                } else {
                    minArea = area;
                    minAreas.clear();
                    minAreas.add( minArea );
                }
            }
            for (Area area : minAreas) {
                area.addCommand(commandData.getAreaRef(), commandData.getCommand());
            }
        }
        if (clearTemplateCells) {
            userAreas.forEach(area -> area.clearCells());
        }
        return userAreas;
    }
}
