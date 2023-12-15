package org.jxls.builder.xls;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.jxls.area.Area;
import org.jxls.area.CommandData;
import org.jxls.area.XlsArea;
import org.jxls.builder.AreaBuilder;
import org.jxls.command.Command;
import org.jxls.command.EachCommand;
import org.jxls.command.GridCommand;
import org.jxls.command.IfCommand;
import org.jxls.command.UpdateCellCommand;
import org.jxls.common.AreaRef;
import org.jxls.common.CellData;
import org.jxls.common.CellRef;
import org.jxls.transform.Transformer;
import org.jxls.util.LiteralsExtractor;
import org.jxls.util.Util;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Builds {@link org.jxls.area.XlsArea} from Excel comments in the Excel template
 * 
 * <h2>Command syntax</h2>
 * <p>A command is specified in a cell comment like the following</p>
 * <pre> jx:COMMAND_NAME(attr1="value1" attr2="value2" ... attrN="valueN" lastCell="LAST_CELL" areas=["AREA_REF1", "AREA_REF2", ... , "AREA_REFN"])</pre>
 * where<ul>
 * <li>COMMAND_NAME - the name of the command</li>
 *
 * <li>attr1, attr2, ... attrN, value1, value2, ... , valueN - command attributes and their values</li>
 *
 * <li>lastCell, LAST_CELL - attribute name and cell reference value specifying the last cell where this command is placed
 * in the parent area. The first cell is defined by the cell where the comment is defined.
 * If there is no "areas" attribute defined it also defines the single area for this command to operate on.</li>
 *
 * <li>AREA_REF1, AREA_REF2, ... , AREA_REFN - additional area references for this command (if supported by the command)
 * 'areas' attribute is optional and only needed for commands which work with more than one area.
 * If there is only a single area for the command it is usually enough to define just lastCell attribute</li>
 * </ul>
 *
 * <p>Multiple commands can be specified in a single cell comment separated by new lines.
 * In this case the area of the first command will contain the second command and so on.</p>
 *
 * <p>This class defines the following pre-defined mappings between the command names and Command classes:</p>
 * <pre> "jx:each" - {@link org.jxls.command.EachCommand}
 * "jx:if" - {@link IfCommand}
 * "jx:area" - {@link AreaCommand} - for defining the top areas
 * "jx:image" - {@link ImageCommand}
 * "jx:grid" - {@link GridCommand}
 * "jx:updateCell" - {@link UpdateCellCommand}
 * "jx:mergeCells" - {@link MergeCellsCommand}</pre>
 *
 * <p>Custom command classes mapping can be added using addCommandMapping(String commandName, Class clazz) method</p>
 *
 * <h2>Command examples</h2>
 *
 * <pre> jx:if(condition="employee.payment &lt;= 2000", lastCell="F9", areas=["A9:F9","A30:F30"])</pre>
 *
 * <p>Here we define {@link IfCommand} with a condition expression 'employee.payment &lt;= 2000' and first area (if-area) "A9:F9"
 * and second area (else-area) "A30:F30". The command is added to the parent area covering a range from the cell where
 * the comment is placed and to the cell defined in lastCell attribute "F9".</p>
 *
 * <pre> jx:each(items="department.staff", var="employee", lastCell="F9")</pre>
 *
 * <p>Here we define {@link org.jxls.command.EachCommand} with items attribute set to 'department.staff' and var attribute set to 'employee'.
 * The command area is defined from the cell where the comment is defined and till the lastCell "F9"</p>
 *
 * <pre> jx:area(lastCell="G26")</pre>
 *
 * <p>Specifies the top area range with {@link AreaCommand} starting from the cell where the comment is defined and in
 * the cell defined in lastCell("G26").</p>
 *
 * <p>Note: Clearing comments from the cells appears to have some issues in POI so should be used with caution.
 * The easiest approach will be just removing the template sheet.</p>
 *
 * @author Leonid Vysochyn
 */
public class XlsCommentAreaBuilder implements AreaBuilder {
    private static Logger logger = LoggerFactory.getLogger(XlsCommentAreaBuilder.class);

    public static final String COMMAND_PREFIX = "jx:";
    private static final String ATTR_PREFIX = "(";
    private static final String ATTR_SUFFIX = ")";
    public static final String LINE_SEPARATOR = "__LINE_SEPARATOR__";
    /**
     * Feature toggle for the multi-line SQL feature (#79). By default turned off for 2.10.0 (incubating). By default turned on starting with version 2.11.0.
     */
    public static boolean MULTI_LINE_SQL_FEATURE = false;
    /*
     * In addition to normal (straight) single and double quotes, this regex
     * includes the following commonly occurring quote-like characters (some
     * of which have been observed in recent versions of LibreOffice):
     *
     * U+201C - LEFT DOUBLE QUOTATION MARK
     * U+201D - RIGHT DOUBLE QUOTATION MARK
     * U+201E - DOUBLE LOW-9 QUOTATION MARK
     * U+201F - DOUBLE HIGH-REVERSED-9 QUOTATION MARK
     * U+2033 - DOUBLE PRIME
     * U+2036 - REVERSED DOUBLE PRIME
     * U+2018 - LEFT SINGLE QUOTATION MARK
     * U+2019 - RIGHT SINGLE QUOTATION MARK
     * U+201A - SINGLE LOW-9 QUOTATION MARK
     * U+201B - SINGLE HIGH-REVERSED-9 QUOTATION MARK
     * U+2032 - PRIME
     * U+2035 - REVERSED PRIME
     */
    private static final String ATTR_REGEX = "\\s*\\w+\\s*=\\s*([\"|'\u201C\u201D\u201E\u201F\u2033\u2036\u2018\u2019\u201A\u201B\u2032\u2035])(?:(?!\\1).)*\\1";
    private static final Pattern ATTR_REGEX_PATTERN = Pattern.compile(ATTR_REGEX);
    private static final String AREAS_ATTR_REGEX = "areas\\s*=\\s*\\[[^]]*]";
    private static final Pattern AREAS_ATTR_REGEX_PATTERN = Pattern.compile(AREAS_ATTR_REGEX);
    
    private static Map<String, Class<? extends Command>> commandMap = new ConcurrentHashMap<>();
    private static final String LAST_CELL_ATTR_NAME = "lastCell";

    static {
        commandMap.put(EachCommand.COMMAND_NAME, EachCommand.class);
        commandMap.put(IfCommand.COMMAND_NAME, IfCommand.class);
        commandMap.put(AreaCommand.COMMAND_NAME, AreaCommand.class);
        commandMap.put(GridCommand.COMMAND_NAME, GridCommand.class);
        commandMap.put(UpdateCellCommand.COMMAND_NAME, UpdateCellCommand.class);
    }

    public static void addCommandMapping(String commandName, Class<? extends Command> clazz) {
        commandMap.put(commandName, clazz);
    }

	public static void removeCommandMapping(String commandName) {
		commandMap.remove(commandName);
	}
	
	public static Class<? extends Command> getCommandClass(String commandName) {
		return commandMap.get(commandName);
	}

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
        List<Area> userAreas = new ArrayList<Area>();
        List<CellData> commentedCells = transformer.getCommentedCells();
        List<CommandData> allCommands = new ArrayList<CommandData>();
        List<Area> allAreas = new ArrayList<Area>();
        for (CellData cellData : commentedCells) {
            String comment = cellData.getCellComment();
            List<CommandData> commandDatas = buildCommands(transformer, cellData, comment);
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
        for (int i = 0; i < allCommands.size(); i++) {
            CommandData commandData = allCommands.get(i);
            AreaRef commandAreaRef = commandData.getAreaRef();
            List<Area> commandAreas = commandData.getCommand().getAreaList();
            Area minArea = null;
            List<Area> minAreas = new ArrayList<Area>();
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
    
    private List<CommandData> buildCommands(Transformer transformer, CellData cellData, String text) {
        List<CommandData> commandDatas = new ArrayList<CommandData>();
        List<String> commentLines;
        if (MULTI_LINE_SQL_FEATURE) {
            commentLines = new LiteralsExtractor().extract(text);
        } else {
            commentLines = Arrays.asList(text.split("\\n"));
        }
        for (String commentLine : commentLines) {
            String line = commentLine.trim();
            if (MULTI_LINE_SQL_FEATURE) {
                line = line.replace("\r\n", LINE_SEPARATOR)
                        .replace("\r", LINE_SEPARATOR)
                        .replace("\n", LINE_SEPARATOR);
            }
            if (!isCommandString(line)) {
                continue;
            }
            int nameEndIndex = line.indexOf(ATTR_PREFIX, COMMAND_PREFIX.length());
            if (nameEndIndex < 0) {
                String errMsg = "Failed to parse command line [" + line + "]. Expected '" + ATTR_PREFIX + "' symbol.";
                throw new IllegalStateException(errMsg);
            }
            String commandName = line.substring(COMMAND_PREFIX.length(), nameEndIndex).trim();
            Map<String, String> attrMap = buildAttrMap(line, nameEndIndex);
            CommandData commandData = createCommandData(cellData, commandName, attrMap);
            if (commandData != null) {
                commandDatas.add(commandData);
                List<Area> areas = buildAreas(transformer, cellData, line);
                for (Area area : areas) {
                    commandData.getCommand().addArea(area);
                }
                if (areas.isEmpty()) {
                    Area area = new XlsArea(commandData.getAreaRef(), transformer);
                    commandData.getCommand().addArea(area);
                }
            }
        }
        return commandDatas;
    }

    public static boolean isCommandString(String str) {
        return str.startsWith(COMMAND_PREFIX) && !str.startsWith(CellData.JX_PARAMS_PREFIX);
    }

    private List<Area> buildAreas(Transformer transformer, CellData cellData, String commandLine) {
        List<Area> areas = new ArrayList<Area>();
        Matcher areasAttrMatcher = AREAS_ATTR_REGEX_PATTERN.matcher(commandLine);
        if (areasAttrMatcher.find()) {
            String areasAttr = areasAttrMatcher.group();
            List<AreaRef> areaRefs = extractAreaRefs(cellData, areasAttr);
            for (AreaRef areaRef : areaRefs) {
                Area area = new XlsArea(areaRef, transformer);
                areas.add(area);
            }
        }
        return areas;
    }

    private List<AreaRef> extractAreaRefs(CellData cellData, String areasAttr) {
        List<AreaRef> areaRefs = new ArrayList<AreaRef>();
        Matcher areaRefMatcher = Util.regexAreaRefPattern.matcher(areasAttr);
        while (areaRefMatcher.find()) {
            String areaRefName = areaRefMatcher.group();
            AreaRef areaRef = new AreaRef(areaRefName);
            if (areaRef.getSheetName() == null || areaRef.getSheetName().trim().length() == 0) {
                areaRef.getFirstCellRef().setSheetName(cellData.getSheetName());
            }
            areaRefs.add(areaRef);
        }
        return areaRefs;
    }

    private Map<String, String> buildAttrMap(String commandLine, int nameEndIndex) {
        int paramsEndIndex = commandLine.lastIndexOf(ATTR_SUFFIX);
        if (paramsEndIndex < 0) {
            String errMsg = "Failed to parse command line [" + commandLine + "]. Expected '" + ATTR_SUFFIX + "' symbol.";
            logger.error(errMsg);
            throw new IllegalArgumentException(errMsg);
        }
        String attrString = commandLine.substring(nameEndIndex + 1, paramsEndIndex).trim();
        return parseCommandAttributes(attrString);
    }

    private CommandData createCommandData(CellData cellData, String commandName, Map<String, String> attrMap) {
        Class<? extends Command> clazz = commandMap.get(commandName);
        if (clazz == null) {
            logger.warn("Failed to find Command class mapped to command name '" + commandName + "'");
            return null;
        }
        try {
            Command command = clazz.getDeclaredConstructor().newInstance();
            for (Map.Entry<String, String> attr : attrMap.entrySet()) {
                if (!attr.getKey().equals(LAST_CELL_ATTR_NAME)) {
                    Util.setObjectProperty(command, attr.getKey(), attr.getValue(), true);
                }
            }
            String lastCellRef = attrMap.get(LAST_CELL_ATTR_NAME);
            if (lastCellRef == null) {
                logger.warn("Failed to find last cell ref attribute '" + LAST_CELL_ATTR_NAME + "' for command '"
                        + commandName + "' in cell " + cellData.getCellRef());
                return null;
            }
            CellRef lastCell = new CellRef(lastCellRef);
            if (lastCell.getSheetName() == null || lastCell.getSheetName().trim().length() == 0) {
                lastCell.setSheetName(cellData.getSheetName());
            }
            return new CommandData(new AreaRef(cellData.getCellRef(), lastCell), command);
        } catch (Exception e) {
            logger.warn("Failed to instantiate command class '" + clazz.getName() + "' mapped to command name '" + commandName + "'", e);
            return null;
        }
    }

    private Map<String, String> parseCommandAttributes(String attrString) {
        Map<String, String> attrMap = new LinkedHashMap<String, String>();
        Matcher attrMatcher = ATTR_REGEX_PATTERN.matcher(attrString);
        while (attrMatcher.find()) {
            String attrData = attrMatcher.group();
            int attrNameEndIndex = attrData.indexOf("=");
            String attrName = attrData.substring(0, attrNameEndIndex).trim();
            String attrValuePart = attrData.substring(attrNameEndIndex + 1).trim();
            String attrValue = attrValuePart.substring(1, attrValuePart.length() - 1);
            attrMap.put(attrName, attrValue);
        }
        return attrMap;
    }
}
