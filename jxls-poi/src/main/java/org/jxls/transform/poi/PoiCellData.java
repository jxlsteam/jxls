package org.jxls.transform.poi;

import java.util.Date;
import java.util.Map;

import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.util.Util;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCell;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Cell data wrapper for POI cell
 * 
 * @author Leonid Vysochyn
 */
public class PoiCellData extends org.jxls.common.CellData {
    private static Logger logger = LoggerFactory.getLogger(PoiCellData.class);

    private PoiRowData poiRowData;
    private RichTextString richTextString;
    private CellStyle cellStyle;
    private Hyperlink hyperlink;
    private Comment comment;
    private String commentAuthor;
    private Cell cell;

    public PoiCellData(CellRef cellRef) {
        super(cellRef);
    }

    public PoiCellData(CellRef cellRef, Cell cell) {
        super(cellRef);
        this.cell = cell;
    }

    public static PoiCellData createCellData(PoiRowData poiRowData, CellRef cellRef, Cell cell){
        PoiCellData cellData = new PoiCellData(cellRef, cell);
        cellData.poiRowData = poiRowData;
        cellData.readCell(cell);
        cellData.updateFormulaValue();
        return cellData;
    }

    public void readCell(Cell cell){
        readCellGeneralInfo(cell);
        readCellContents(cell);
        readCellStyle(cell);
    }

    private void readCellGeneralInfo(Cell cell) {
        hyperlink = cell.getHyperlink();
        try {
            comment = cell.getCellComment();
        } catch (Exception e) {
            logger.error("Failed to read cell comment at " + new CellReference(cell).formatAsString(), e);
            return;
        }
        if (comment != null) {
            commentAuthor = comment.getAuthor();
        }
        if (comment != null && comment.getString() != null && comment.getString().getString() != null) {
            String commentString = comment.getString().getString();
            String[] commentLines = commentString.split("\\n");
            for (String commentLine : commentLines) {
                if (isJxlsParamsComment(commentLine)) {
                    processJxlsParams(commentLine);
                    comment = null;
                    return;
                }
            }
            setCellComment(commentString);
        }
    }

    public CellStyle getCellStyle() {
        return cellStyle;
    }

    public void setCellStyle(CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }

    private void readCellContents(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                richTextString = cell.getRichStringCellValue();
                cellValue = richTextString.getString();
                cellType = CellType.STRING;
                break;
            case BOOLEAN:
                cellValue = cell.getBooleanCellValue();
                cellType = CellType.BOOLEAN;
                break;
            case NUMERIC:
                readNumericCellContents(cell);
                break;
            case FORMULA:
                formula = cell.getCellFormula();
                cellValue = formula;
                cellType = CellType.FORMULA;
                break;
            case ERROR:
                cellValue = cell.getErrorCellValue();
                cellType = CellType.ERROR;
                break;
            case BLANK:
            case _NONE:
                cellValue = null;
                cellType = CellType.BLANK;
                break;
        }
        evaluationResult = cellValue;
    }

    private void readNumericCellContents(Cell cell) {
        if (DateUtil.isCellDateFormatted(cell)) {
            cellValue = cell.getDateCellValue();
            cellType = CellType.DATE;
        } else {
            cellValue = cell.getNumericCellValue();
            cellType = CellType.NUMBER;
        }
    }

    private void readCellStyle(Cell cell) {
        cellStyle = cell.getCellStyle();
    }

    public void writeToCell(Cell cell, Context context, PoiTransformer transformer) {
        evaluate(context);
        if (evaluationResult instanceof WritableCellValue) {
            cell.setCellStyle(cellStyle);
            ((WritableCellValue) evaluationResult).writeToCell(cell, context);
        } else {
            updateCellGeneralInfo(cell);
            updateCellContents(cell);
            CellStyle targetCellStyle = cellStyle;
            if (context.getConfig().isIgnoreSourceCellStyle()) {
                CellStyle dataFormatCellStyle = findCellStyle(evaluationResult, context.getConfig().getCellStyleMap(), transformer);
                if (dataFormatCellStyle != null) {
                    targetCellStyle = dataFormatCellStyle;
                }
            }
            updateCellStyle(cell, targetCellStyle);
            poiRowData.getPoiSheetData().updateConditionalFormatting(this, cell);
        }
    }

    private CellStyle findCellStyle(Object evaluationResult, Map<String, String> cellStyleMap, PoiTransformer transformer) {
        if (evaluationResult == null || cellStyleMap == null) {
            return null;
        }
        String cellName = cellStyleMap.get(evaluationResult.getClass().getSimpleName());
        if (cellName == null) {
            return null;
        }
        Sheet sheet = cell.getSheet();
        CellRef cellRef = new CellRef(cellName);
        if (cellRef.getSheetName() == null) {
            cellRef.setSheetName(sheet.getSheetName());
        }
        return transformer.getCellStyle(cellRef);
    }

    private void updateCellGeneralInfo(Cell cell) {
        if (targetCellType != CellType.FORMULA) {
            cell.setCellType(getPoiCellType(targetCellType));
        }
        if (hyperlink != null) {
            cell.setHyperlink(hyperlink);
        }
        if (comment != null && !PoiUtil.isJxComment(getCellComment())) {
            PoiUtil.setCellComment(cell, getCellComment(), commentAuthor, null);
        }
    }

    static org.apache.poi.ss.usermodel.CellType getPoiCellType(CellType cellType) {
        if (cellType == null) {
            return org.apache.poi.ss.usermodel.CellType.BLANK;
        }
        switch (cellType) {
            case STRING:  return org.apache.poi.ss.usermodel.CellType.STRING;
            case BOOLEAN: return org.apache.poi.ss.usermodel.CellType.BOOLEAN;
            case NUMBER:
            case DATE:    return org.apache.poi.ss.usermodel.CellType.NUMERIC;
            case FORMULA: return org.apache.poi.ss.usermodel.CellType.FORMULA;
            case ERROR:   return org.apache.poi.ss.usermodel.CellType.ERROR;
            case BLANK:   return org.apache.poi.ss.usermodel.CellType.BLANK;
            default:      return org.apache.poi.ss.usermodel.CellType.BLANK;
        }
    }

    private void updateCellContents(Cell cell) {
        switch (targetCellType) {
            case STRING:
                updateStringCellContents(cell);
                break;
            case BOOLEAN:
                cell.setCellValue((Boolean) evaluationResult);
                break;
            case DATE:
                cell.setCellValue((Date) evaluationResult);
                break;
            case NUMBER:
                cell.setCellValue(((Number) evaluationResult).doubleValue());
                break;
            case FORMULA:
                updateFormulaCellContents(cell);
                break;
            case ERROR:
                cell.setCellErrorValue((Byte) evaluationResult);
                break;
            case BLANK:
                cell.setCellType(org.apache.poi.ss.usermodel.CellType.BLANK);
                break;
        }
    }

    private void updateStringCellContents(Cell cell) {
        if (evaluationResult instanceof byte[]) {
            return;
        }
        String result = evaluationResult != null ? evaluationResult.toString() : "";
        if (cellValue != null && cellValue.equals(result)) {
            cell.setCellValue(richTextString);
        } else {
            cell.setCellValue(result);
        }
    }

    private void updateFormulaCellContents(Cell cell) {
        try {
            if (Util.formulaContainsJointedCellRef((String) evaluationResult)) {
                cell.setCellValue((String) evaluationResult);
            } else {
                cell.setCellFormula((String) evaluationResult);
                clearCellValue(cell); // This call is especially important for streaming.
            }
        } catch (FormulaParseException e) {
            try {
                String formulaString = evaluationResult.toString();
                logger.error("Failed to set cell formula " + formulaString + " for cell " + this.toString(), e);
                cell.setCellType(org.apache.poi.ss.usermodel.CellType.STRING);
                cell.setCellValue(formulaString);
            } catch (Exception ex) {
                logger.warn("Failed to convert formula to string for cell " + this.toString());
            }
        }
    }

    // protected so any user can change this piece of code
    protected void clearCellValue(org.apache.poi.ss.usermodel.Cell poiCell) {
        if (poiCell instanceof XSSFCell) {
            CTCell cell = ((XSSFCell) poiCell).getCTCell(); // POI internal access, but there's no other way
            // Now do the XSSFCell.setFormula code that was done before POI commit https://github.com/apache/poi/commit/1253a29
            // After setting the formula in attribute f we clear the value attribute v if set. This causes a recalculation
            // and prevents wrong formula results.
            if (cell.isSetV()) {
                cell.unsetV();
            }
        }
    }

    private void updateCellStyle(Cell cell, CellStyle cellStyle) {
        cell.setCellStyle(cellStyle);
    }
}
