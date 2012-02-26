package com.jxls.writer.builder.xls;

import com.jxls.writer.area.Area;
import com.jxls.writer.area.CommandData;
import com.jxls.writer.area.XlsArea;
import com.jxls.writer.builder.AreaBuilder;
import com.jxls.writer.command.Command;
import com.jxls.writer.common.CellData;
import com.jxls.writer.transform.Transformer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

/**
 * @author Leonid Vysochyn
 */
public class XlsCommentAreaBuilder implements AreaBuilder {
    static Logger logger = LoggerFactory.getLogger(XlsCommentAreaBuilder.class);

    private static final String COMMAND_PREFIX = "jx:";
    private static final String ATTR_PREFIX = "(";
    private static final String ATTR_SUFFIX = ")";
    private static final String ATTR_REGEX = "\\s*\\w+\\s*=\\s*([\"|'])(?:(?!\\1).)*\\1";

    Transformer transformer;

    public XlsCommentAreaBuilder(Transformer transformer) {
        this.transformer = transformer;
    }

    public List<Area> build() {
        List<Area> areas = new ArrayList<Area>();
        List<CellData> commentedCells = transformer.getCommentedCells();
        Area currentArea = null;
        for (CellData cellData : commentedCells) {
            String comment = cellData.getCellComment();
            List<CommandData> commandDatas = buildCommands(cellData, comment);
            if( currentArea == null ){
                // todo
            }
            for (CommandData commandData : commandDatas) {
                
            }
        }
        return areas;
    }

    private List<CommandData> buildCommands(CellData cellData, String text) {
        String[] commandLines = text.split("\\n");
        List<CommandData> commands = new ArrayList<CommandData>();
        for (int i = 0; i < commandLines.length; i++) {
            String commandLine = commandLines[i].trim();
            if(commandLine.startsWith(COMMAND_PREFIX)){
                int nameEndIndex = commandLine.indexOf(ATTR_PREFIX, COMMAND_PREFIX.length());
                if( nameEndIndex < 0 ){
                    String errMsg = "Failed to parse command line [" + commandLine + "]. Expected '" + ATTR_PREFIX + "' symbol.";
                    logger.error(errMsg);
                    throw new IllegalStateException(errMsg);
                }
                String commandName = commandLine.substring(COMMAND_PREFIX.length(), nameEndIndex).trim();
                int paramsEndIndex = commandLine.lastIndexOf(ATTR_SUFFIX);
                if(paramsEndIndex < 0 ){
                    String errMsg = "Failed to parse command line [" + commandLine + "]. Expected '" + ATTR_SUFFIX + "' symbol.";
                    logger.error(errMsg);
                    throw new IllegalArgumentException(errMsg);
                }
                String attrString = commandLine.substring(nameEndIndex + 1, paramsEndIndex).trim();
                Map<String, String> attrMap = parseCommandAttributes(attrString);
                CommandData commandData = createCommandData(commandName, attrMap);
                commands.add(commandData);
            }else{
                logger.info("Command line does not start with command prefix '" + COMMAND_PREFIX + "'. Skipping it.");
            }
        }
        return commands;
    }

    private CommandData createCommandData(String commandName, Map<String, String> attrMap) {
        return null;
    }

    private Map<String, String> parseCommandAttributes(String attrString) {
        Map<String,String> attrMap = new LinkedHashMap<String, String>();
        String[] attrDatas = attrString.split(ATTR_REGEX);
        for (int i = 0; i < attrDatas.length; i++) {
            String attrData = attrDatas[i].trim();
            int attrNameEndIndex = attrData.indexOf("=");
            String attrName = attrData.substring(0, attrNameEndIndex).trim();
            String attrValuePart = attrData.substring(attrNameEndIndex + 1).trim();
            String attrValue = attrValuePart.substring(1, attrValuePart.length() - 1);
            attrMap.put(attrName, attrValue);
        }
        return attrMap;
    }

}
