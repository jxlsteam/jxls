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
import org.jxls.builder.CommandMappings;
import org.jxls.command.Command;
import org.jxls.command.EachCommand;
import org.jxls.command.GridCommand;
import org.jxls.command.IfCommand;
import org.jxls.command.UpdateCellCommand;
import org.jxls.common.AreaListener;
import org.jxls.common.AreaRef;
import org.jxls.common.CellData;
import org.jxls.common.CellRef;
import org.jxls.common.ObjectPropertyAccess;
import org.jxls.formula.AbstractFormulaProcessor;
import org.jxls.logging.JxlsLogger;
import org.jxls.transform.Transformer;
import org.jxls.util.LiteralsExtractor;

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
 * "jx:grid" - {@link GridCommand}
 * "jx:updateCell" - {@link UpdateCellCommand}</pre>
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
public class XlsCommentAreaBuilder extends AbstractAreaBuilder implements CommandMappings {
    public static final String COMMAND_PREFIX = "jx:";
    private static final String ATTR_PREFIX = "(";
    private static final String ATTR_SUFFIX = ")";
    public static final String LINE_SEPARATOR = "__LINE_SEPARATOR__";
    /** Feature toggle for the multi-line SQL feature (#79) */
    public static boolean MULTI_LINE_SQL_FEATURE = true;
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
    private static final String LAST_CELL_ATTR_NAME = "lastCell";
    private static final String regexSimpleCellRef = "[a-zA-Z]+[0-9]+";
    private static final String regexAreaRef = AbstractFormulaProcessor.regexCellRef + ":" + regexSimpleCellRef;
    private static final Pattern regexAreaRefPattern = Pattern.compile(regexAreaRef);

    private final Map<String, Class<? extends Command>> commandMap = new ConcurrentHashMap<>();

    public XlsCommentAreaBuilder() {
        addCommandMapping(EachCommand.COMMAND_NAME, EachCommand.class);
        addCommandMapping(IfCommand.COMMAND_NAME, IfCommand.class);
        addCommandMapping(AreaCommand.COMMAND_NAME, AreaCommand.class);
        addCommandMapping(GridCommand.COMMAND_NAME, GridCommand.class);
        addCommandMapping(UpdateCellCommand.COMMAND_NAME, UpdateCellCommand.class);
    }

    @Override
    public void addCommandMapping(String commandName, Class<? extends Command> commandClass) {
        commandMap.put(commandName, commandClass);
    }

    @Override
	public void removeCommandMapping(String commandName) {
		commandMap.remove(commandName);
	}
	
    @Override
	public Class<? extends Command> getCommandClass(String commandName) {
		return commandMap.get(commandName);
	}
    
    @Override
    protected List<CommandData> buildCommands(Transformer transformer, CellData cellData, String text) {
        List<CommandData> commandDatas = new ArrayList<>();
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
                throw new JxlsCommentException("Failed to parse command line [" + line + "]. Expected '" + ATTR_PREFIX + "' symbol.");
            }
            String commandName = line.substring(COMMAND_PREFIX.length(), nameEndIndex).trim();
            Map<String, String> attrMap = buildAttrMap(line, nameEndIndex);
            CommandData commandData = createCommandData(cellData, commandName, attrMap, transformer.getLogger());
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
        List<Area> areas = new ArrayList<>();
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
        List<AreaRef> areaRefs = new ArrayList<>();
        Matcher areaRefMatcher = regexAreaRefPattern.matcher(areasAttr);
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
            throw new JxlsCommentException("Failed to parse command line '" + commandLine + "'. Expected '" + ATTR_SUFFIX + "' symbol.");
        }
        String attrString = commandLine.substring(nameEndIndex + 1, paramsEndIndex).trim();
        return parseCommandAttributes(attrString);
    }

    private CommandData createCommandData(CellData cellData, String commandName, Map<String, String> attrMap, JxlsLogger logger) {
        Class<? extends Command> clazz = getCommandClass(commandName);
        if (clazz == null) {
            throw new JxlsCommentException("Failed to find Command class mapped to command name '" + commandName + "'");
        }
        try {
            Command command = clazz.getDeclaredConstructor().newInstance();
            for (Map.Entry<String, String> attr : attrMap.entrySet()) {
                if (!attr.getKey().equals(LAST_CELL_ATTR_NAME)) {
                    ObjectPropertyAccess.setObjectProperty(command, attr.getKey(), attr.getValue(), logger);
                }
            }
            String lastCellRef = attrMap.get(LAST_CELL_ATTR_NAME);
            if (lastCellRef == null) {
                throw new JxlsCommentException("Failed to find attribute '" + LAST_CELL_ATTR_NAME + "' for command '"
                        + commandName + "' in cell " + cellData.getCellRef());
            }
            CellRef lastCell = new CellRef(lastCellRef);
            if (lastCell.getSheetName() == null || lastCell.getSheetName().trim().length() == 0) {
                lastCell.setSheetName(cellData.getSheetName());
            }
            return new CommandData(new AreaRef(cellData.getCellRef(), lastCell), command);
        } catch (ReflectiveOperationException | IllegalArgumentException | SecurityException e) {
            throw new JxlsCommentException("Failed to instantiate command class '" + clazz.getName()
            	+ "' mapped to command name '" + commandName + "'", e);
		}
    }

    private Map<String, String> parseCommandAttributes(String attrString) {
        Map<String, String> attrMap = new LinkedHashMap<>();
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
    
    /**
     * Method for adding an AreaListener to an area given by an AreaRef
     * @param areaListener to be added AreaListener
     * @param areaRef area where the AreaListener has to be added
     * @param areas all areas to search for
     */
    protected void addAreaListener(AreaListener areaListener, AreaRef areaRef, List<Area> areas) {
        for (Area area : areas) {
            if (areaRef.equals(area.getAreaRef())) {
                area.addAreaListener(areaListener);
                return;
            }
            for (CommandData command : area.getCommandDataList()) {
                addAreaListener(areaListener, areaRef, command.getCommand().getAreaList()); // recursive
            }
        }
    }
}
