package org.jxls.transform.poi;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.ConditionalFormatting;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jxls.common.AreaRef;
import org.jxls.common.CellData;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.ImageType;
import org.jxls.common.RowData;
import org.jxls.common.SheetData;
import org.jxls.common.Size;
import org.jxls.transform.AbstractTransformer;
import org.jxls.util.CannotOpenWorkbookException;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCell;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * POI implementation of {@link org.jxls.transform.Transformer} interface
 *
 * @author Leonid Vysochyn
 */
public class PoiTransformer extends AbstractTransformer {
    private static final int MAX_COLUMN_TO_READ_COMMENT = 50;
    public static final String POI_CONTEXT_KEY = "util";

    private static Logger logger = LoggerFactory.getLogger(PoiTransformer.class);

    private Workbook workbook;
    private OutputStream outputStream;
    private InputStream inputStream;
    private Integer lastCommentedColumn = MAX_COLUMN_TO_READ_COMMENT;
    private final boolean isSXSSF;

    /**
     * The cell style is lost after the merge, the following operation restores the merged cell
     * to the style of the first cell before the merge.
     */
    private CellStyle cellStyle;

    /**
     * No streaming
     * @param workbook
     */
    private PoiTransformer(Workbook workbook) {
        this(workbook, false);
    }

    /**
     * @param workbook source workbook to transform
     * @param streaming false: without streaming, true: with streaming (with default parameter values)
     */
    public PoiTransformer(Workbook workbook, boolean streaming) {
        this(workbook, streaming, SXSSFWorkbook.DEFAULT_WINDOW_SIZE, false, false);
    }

    /**
     * @param workbook source workbook to transform
     * @param streaming flag to set if SXSSF stream support is enabled
     * @param rowAccessWindowSize only used if streaming is true
     * @param compressTmpFiles only used if streaming is true
     * @param useSharedStringsTable only used if streaming is true
     */
    public PoiTransformer(Workbook workbook, boolean streaming, int rowAccessWindowSize, boolean compressTmpFiles, boolean useSharedStringsTable) {
        this.workbook = workbook;
        isSXSSF = streaming;
        readCellData();
        if (isSXSSF) {
            if (this.workbook instanceof XSSFWorkbook) {
                this.workbook = new SXSSFWorkbook((XSSFWorkbook) this.workbook, rowAccessWindowSize, compressTmpFiles, useSharedStringsTable);
            } else {
                throw new IllegalArgumentException("Failed to create POI Transformer using SXSSF API as the input workbook is not XSSFWorkbook");
            }
        }
    }
    
    protected boolean isStreaming() {
        return isSXSSF;
    }
    
    public void setInputStream(InputStream is) {
        inputStream = is;
    }

    /**
     * Creates transformer from an input stream template and output stream
     * @param is input stream to read the Excel template file. Format can be XLSX (recommended) or XLS.
     * @param os output stream to write the Excel file. Must be the same format.
     * @return {@link PoiTransformer} instance
     */
    public static PoiTransformer createTransformer(InputStream is, OutputStream os) {
        PoiTransformer transformer = createTransformer(is);
        transformer.setOutputStream(os);
        transformer.setInputStream(is);
        return transformer;
    }

    /**
     * Creates transformer instance for given input stream
     * @param is input stream for the Excel template file. Format can be XLSX (recommended) or XLS.
     * @return transformer instance reading the template from the passed input stream
     * @throws CannotOpenWorkbookException if an error occurs during opening the Excel workbook
     */
    public static PoiTransformer createTransformer(InputStream is) {
        Workbook workbook;
        try {
            workbook = WorkbookFactory.create(is);
        } catch (Exception e) {
            throw new CannotOpenWorkbookException(e);
        }
        return createTransformer(workbook);
    }

    /**
     * Creates transformer instance from a {@link Workbook} instance
     * @param workbook Excel template
     * @return transformer instance with the given workbook as template
     */
    public static PoiTransformer createTransformer(Workbook workbook) {
        return new PoiTransformer(workbook);
    }

    /**
     * Creates transformer for given workbook. Streaming will be used.
     * @param workbook Excel template. Format must be XLSX.
     * @return transformer instance with the given workbook as template
     */
    public static PoiTransformer createSxssfTransformer(Workbook workbook) {
        return createSxssfTransformer(workbook, SXSSFWorkbook.DEFAULT_WINDOW_SIZE, false);
    }

    /**
     * Creates transformer for given workbook and streaming parameters. Streaming will be used.
     * @param workbook Excel template. Format must be XLSX.
     * @param rowAccessWindowSize -
     * @param compressTmpFiles -
     * @return transformer instance with the given workbook as template
     */
    public static PoiTransformer createSxssfTransformer(Workbook workbook, int rowAccessWindowSize, boolean compressTmpFiles) {
        return createSxssfTransformer(workbook, rowAccessWindowSize, compressTmpFiles, false);
    }

    /**
     * Creates transformer for given workbook and streaming parameters. Streaming will be used.
     * @param workbook Excel template. Format must be XLSX.
     * @param rowAccessWindowSize -
     * @param compressTmpFiles -
     * @param useSharedStringsTable -
     * @return transformer instance with the given workbook as template
     */
    public static PoiTransformer createSxssfTransformer(Workbook workbook, int rowAccessWindowSize, boolean compressTmpFiles, boolean useSharedStringsTable) {
        return new PoiTransformer(workbook, true, rowAccessWindowSize, compressTmpFiles, useSharedStringsTable);
    }

    public static Context createInitialContext() {
        Context context = new Context();
        context.putVar(POI_CONTEXT_KEY, new PoiUtil());
        return context;
    }

    @Override
    public boolean isForwardOnly() {
        return isStreaming();
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public Integer getLastCommentedColumn() {
        return lastCommentedColumn;
    }

    public void setLastCommentedColumn(Integer lastCommentedColumn) {
        this.lastCommentedColumn = lastCommentedColumn;
    }

    private void readCellData() {
        int numberOfSheets = workbook.getNumberOfSheets();
        for (int i = 0; i < numberOfSheets; i++) {
            Sheet sheet = workbook.getSheetAt(i);
            SheetData sheetData = PoiSheetData.createSheetData(sheet, this);
            sheetMap.put(sheetData.getSheetName(), sheetData);
        }
    }

    @Override
    public void transform(CellRef srcCellRef, CellRef targetCellRef, Context context, boolean updateRowHeightFlag) {
        CellData cellData = isTransformable(srcCellRef, targetCellRef);
        if (cellData == null) {
            return;
        }
        Sheet destSheet = workbook.getSheet(targetCellRef.getSheetName());
        if (destSheet == null) {
            destSheet = workbook.createSheet(targetCellRef.getSheetName());
            PoiUtil.copySheetProperties(workbook.getSheet(srcCellRef.getSheetName()), destSheet);
        }
        Row destRow = destSheet.getRow(targetCellRef.getRow());
        if (destRow == null) {
            destRow = destSheet.createRow(targetCellRef.getRow());
        }
        transformCell(srcCellRef, targetCellRef, context, updateRowHeightFlag, cellData, destSheet, destRow);
    }
    
    protected CellData isTransformable(CellRef srcCellRef, CellRef targetCellRef) {
        CellData cellData = getCellData(srcCellRef);
        if (cellData != null) {
            if (targetCellRef == null || targetCellRef.getSheetName() == null) {
                logger.info("Target cellRef is null or has empty sheet name, cellRef=" + targetCellRef);
                return null; // do not transform
            }
        }
        return cellData;
    }

    protected void transformCell(CellRef srcCellRef, CellRef targetCellRef, Context context,
            boolean updateRowHeightFlag, CellData cellData, Sheet destSheet, Row destRow) {
        SheetData sheetData = sheetMap.get(srcCellRef.getSheetName());
        if (!isIgnoreColumnProps()) {
            destSheet.setColumnWidth(targetCellRef.getCol(), sheetData.getColumnWidth(srcCellRef.getCol()));
        }
        if (updateRowHeightFlag && !isIgnoreRowProps()) {
            destRow.setHeight((short) sheetData.getRowData(srcCellRef.getRow()).getHeight());
        }
        org.apache.poi.ss.usermodel.Cell destCell = destRow.getCell(targetCellRef.getCol());
        if (destCell == null) {
            destCell = destRow.createCell(targetCellRef.getCol());
        }
        try {
            // conditional formatting
            destCell.setCellType(CellType.BLANK);
            ((PoiCellData) cellData).writeToCell(destCell, context, this);
            copyMergedRegions(cellData, targetCellRef);
        } catch (Exception e) {
            logger.error("Failed to write a cell with {} and context keys {}", cellData, context.toMap().keySet(), e);
        }
    }

    @Override
    public void resetArea(AreaRef areaRef) {
        removeMergedRegions(areaRef);
        removeConditionalFormatting(areaRef);
    }

    private void removeMergedRegions(AreaRef areaRef) {
        Sheet destSheet = workbook.getSheet(areaRef.getSheetName());
        int numMergedRegions = destSheet.getNumMergedRegions();
        for (int i = numMergedRegions; i > 0; i--) {
            destSheet.removeMergedRegion(i - 1);
        }
    }

    // this method updates conditional formatting ranges only when the range is inside the passed areaRef
    private void removeConditionalFormatting(AreaRef areaRef) {
        Sheet destSheet = workbook.getSheet(areaRef.getSheetName());
        CellRangeAddress areaRange = CellRangeAddress.valueOf(areaRef.toString());
        SheetConditionalFormatting sheetConditionalFormatting = destSheet.getSheetConditionalFormatting();
        int numConditionalFormattings = sheetConditionalFormatting.getNumConditionalFormattings();
        for (int index = 0; index < numConditionalFormattings; index++) {
            ConditionalFormatting conditionalFormatting = sheetConditionalFormatting.getConditionalFormattingAt(index);
            CellRangeAddress[] ranges = conditionalFormatting.getFormattingRanges();
            List<CellRangeAddress> newRanges = new ArrayList<>();
            for (CellRangeAddress range : ranges) {
                if (!areaRange.isInRange(range.getFirstRow(), range.getFirstColumn()) || !areaRange.isInRange(range.getLastRow(), range.getLastColumn())) {
                    newRanges.add(range);
                }
            }
            conditionalFormatting.setFormattingRanges(newRanges.toArray(new CellRangeAddress[] {}));
        }
    }

    protected final void copyMergedRegions(CellData sourceCellData, CellRef destCell) {
        if (sourceCellData.getSheetName() == null) {
            throw new IllegalArgumentException("Sheet name is null in copyMergedRegions");
        }
        PoiSheetData sheetData = (PoiSheetData) sheetMap.get(sourceCellData.getSheetName());
        CellRangeAddress cellMergedRegion = null;
        for (CellRangeAddress mergedRegion : sheetData.getMergedRegions()) {
            if (mergedRegion.getFirstRow() == sourceCellData.getRow() && mergedRegion.getFirstColumn() == sourceCellData.getCol()) {
                cellMergedRegion = mergedRegion;
                break;
            }
        }
        if (cellMergedRegion != null) {
            findAndRemoveExistingCellRegion(destCell);
            Sheet destSheet = workbook.getSheet(destCell.getSheetName());
            destSheet.addMergedRegion(new CellRangeAddress(destCell.getRow(), destCell.getRow() + cellMergedRegion.getLastRow() - cellMergedRegion.getFirstRow(),
                    destCell.getCol(), destCell.getCol() + cellMergedRegion.getLastColumn() - cellMergedRegion.getFirstColumn()));
        }
    }

    protected final void findAndRemoveExistingCellRegion(CellRef cellRef) {
        Sheet destSheet = workbook.getSheet(cellRef.getSheetName());
        int numMergedRegions = destSheet.getNumMergedRegions();
        for (int i = 0; i < numMergedRegions; i++) {
            CellRangeAddress mergedRegion = destSheet.getMergedRegion(i);
            if (mergedRegion.getFirstRow() <= cellRef.getRow() && mergedRegion.getLastRow() >= cellRef.getRow() &&
                    mergedRegion.getFirstColumn() <= cellRef.getCol() && mergedRegion.getLastColumn() >= cellRef.getCol()) {
                destSheet.removeMergedRegion(i);
                break;
            }
        }
    }

    @Override
    public void setFormula(CellRef cellRef, String formulaString) {
        if (cellRef == null || cellRef.getSheetName() == null) return;
        Sheet sheet = workbook.getSheet(cellRef.getSheetName());
        if (sheet == null) {
            sheet = workbook.createSheet(cellRef.getSheetName());
        }
        Row row = sheet.getRow(cellRef.getRow());
        if (row == null) {
            row = sheet.createRow(cellRef.getRow());
        }
        org.apache.poi.ss.usermodel.Cell poiCell = row.getCell(cellRef.getCol());
        if (poiCell == null) {
            poiCell = row.createCell(cellRef.getCol());
        }
        try {
            poiCell.setCellFormula(formulaString);
            clearCellValue(poiCell);
        } catch (Exception e) {
            logger.error("Failed to set formula = " + formulaString + " into cell = " + cellRef.getCellName(), e);
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

    @Override
    public void clearCell(CellRef cellRef) {
        if (cellRef == null || cellRef.getSheetName() == null) return;
        Sheet sheet = workbook.getSheet(cellRef.getSheetName());
        if (sheet == null) return;
        removeCellComment(sheet, cellRef.getRow(), cellRef.getCol());
        Row row = getRowForClearCell(sheet, cellRef);
        if (row == null) return;
        Cell cell = row.getCell(cellRef.getCol());
        if (cell == null) {
            CellAddress cellAddress = new CellAddress(cellRef.getRow(), cellRef.getCol());
            if (sheet.getCellComment(cellAddress) != null) {
                cell = row.createCell(cellRef.getCol());
                cell.removeCellComment();
            }
            return;
        }
        cell.setCellType(CellType.BLANK);
        cell.setCellStyle(workbook.getCellStyleAt(0));
        if (cell.getCellComment() != null) {
            cell.removeCellComment();
        }
        findAndRemoveExistingCellRegion(cellRef);
    }

    protected Row getRowForClearCell(Sheet sheet, CellRef cellRef) {
        return sheet.getRow(cellRef.getRow());
    }

    protected final void removeCellComment(Sheet sheet, int rowNum, int colNum) {
        Row row = sheet.getRow(rowNum);
        if (row == null) return;
        Cell cell = row.getCell(colNum);
        if (cell == null) return;
        cell.removeCellComment();
    }

    @Override
    public List<CellData> getCommentedCells() {
        List<CellData> commentedCells = new ArrayList<CellData>();
        for (SheetData sheetData : sheetMap.values()) {
            for (RowData rowData : sheetData) {
                if (rowData == null) continue;
                int row = ((PoiRowData) rowData).getRow().getRowNum();
                List<CellData> cellDataList = readCommentsFromSheet(((PoiSheetData) sheetData).getSheet(), row);
                commentedCells.addAll(cellDataList);
            }
        }
        return commentedCells;
    }


    private void addImage(AreaRef areaRef, int imageIdx, Double scaleX, Double scaleY) {
        boolean pictureResizeFlag = scaleX != null && scaleY != null;
        CreationHelper helper = workbook.getCreationHelper();
        Sheet sheet = workbook.getSheet(areaRef.getSheetName());
        if (sheet == null) {
            sheet = workbook.createSheet(areaRef.getSheetName());
        }
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = helper.createClientAnchor();
        anchor.setCol1(areaRef.getFirstCellRef().getCol());
        anchor.setRow1(areaRef.getFirstCellRef().getRow());
        if (pictureResizeFlag) {
            anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_DONT_RESIZE);
            anchor.setCol2(-1);
            anchor.setRow2(-1);
        } else {
            anchor.setCol2(areaRef.getLastCellRef().getCol());
            anchor.setRow2(areaRef.getLastCellRef().getRow());
        }
        Picture picture = drawing.createPicture(anchor, imageIdx);
        if (pictureResizeFlag) {
            picture.resize(scaleX, scaleY);
        }
    }

    @Override
    public void addImage(AreaRef areaRef, byte[] imageBytes, ImageType imageType, Double scaleX, Double scaleY) {
        int poiPictureType = findPoiPictureTypeByImageType(imageType);
        int pictureIdx = workbook.addPicture(imageBytes, poiPictureType);
        addImage(areaRef, pictureIdx, scaleX, scaleY);
    }

    @Override
    public void addImage(AreaRef areaRef, byte[] imageBytes, ImageType imageType) {
        int poiPictureType = findPoiPictureTypeByImageType(imageType);
        int pictureIdx = workbook.addPicture(imageBytes, poiPictureType);
        addImage(areaRef, pictureIdx, null, null);
    }

    @Override
    public void write() throws IOException {
        writeButNotCloseStream();
        outputStream.close();
        dispose();
    }

    @Override
    public void writeButNotCloseStream() throws IOException {
        if (outputStream == null) {
            throw new IllegalStateException("Cannot write a workbook with an uninitialized output stream");
        }
        if (workbook == null) {
            throw new IllegalStateException("Cannot write an uninitialized workbook");
        }
        if (!isStreaming() && isEvaluateFormulas()) {
            workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
        }
        workbook.write(outputStream);
    }

    @Override
    public void dispose() {
        // Note that SXSSF allocates temporary files that you must always clean up explicitly, by calling the dispose method. ( http://poi.apache.org/components/spreadsheet/how-to.html#sxssf )
        try {
            if (workbook instanceof SXSSFWorkbook) {
                ((SXSSFWorkbook) workbook).dispose();
            }
        } catch (Exception e) {
            logger.warn("Error disposing streamed workbook", e);
        }
    }

    private int findPoiPictureTypeByImageType(ImageType imageType) {
        int poiType = -1;
        if (imageType == null) {
            throw new IllegalArgumentException("Image type is undefined");
        }
        switch (imageType) {
            case PNG:
                poiType = Workbook.PICTURE_TYPE_PNG;
                break;
            case JPEG:
                poiType = Workbook.PICTURE_TYPE_JPEG;
                break;
            case EMF:
                poiType = Workbook.PICTURE_TYPE_EMF;
                break;
            case WMF:
                poiType = Workbook.PICTURE_TYPE_WMF;
                break;
            case DIB:
                poiType = Workbook.PICTURE_TYPE_DIB;
                break;
            case PICT:
                poiType = Workbook.PICTURE_TYPE_PICT;
                break;
        }
        return poiType;
    }

    private List<CellData> readCommentsFromSheet(Sheet sheet, int rowNum) {
        List<CellData> commentDataCells = new ArrayList<CellData>();
        for (int i = 0; i <= lastCommentedColumn; i++) {
            CellAddress cellAddress = new CellAddress(rowNum, i);
            Comment comment = sheet.getCellComment(cellAddress);
            if (comment != null && comment.getString() != null) {
                CellData cellData = new CellData(new CellRef(sheet.getSheetName(), rowNum, i));
                cellData.setCellComment(comment.getString().getString());
                commentDataCells.add(cellData);
            }
        }
        return commentDataCells;
    }

    public OutputStream getOutputStream() {
        return outputStream;
    }

    public void setOutputStream(OutputStream outputStream) {
        this.outputStream = outputStream;
    }

    public InputStream getInputStream() {
        return inputStream;
    }

    public CellStyle getCellStyle(CellRef cellRef) {
        SheetData sheetData = sheetMap.get(cellRef.getSheetName());
        PoiCellData cellData = (PoiCellData) sheetData.getRowData(cellRef.getRow()).getCellData(cellRef.getCol());
        return cellData.getCellStyle();
    }

    @Override
    public boolean deleteSheet(String sheetName) {
        if (super.deleteSheet(sheetName)) {
            int sheetIndex = workbook.getSheetIndex(sheetName);
            workbook.removeSheetAt(sheetIndex);
            return true;
        } else {
            logger.warn("Failed to find '{}' worksheet in a sheet map. Skipping the deletion.", sheetName);
            return false;
        }
    }

    @Override
    public void setHidden(String sheetName, boolean hidden) {
        int sheetIndex = workbook.getSheetIndex(sheetName);
        workbook.setSheetHidden(sheetIndex, hidden);
    }

    @Override
    public void updateRowHeight(String srcSheetName, int srcRowNum, String targetSheetName, int targetRowNum) {
        if (isSXSSF) return;
        SheetData sheetData = sheetMap.get(srcSheetName);
        RowData rowData = sheetData.getRowData(srcRowNum);
        Sheet sheet = workbook.getSheet(targetSheetName);
        if (sheet == null) {
            sheet = workbook.createSheet(targetSheetName);
        }
        Row targetRow = sheet.getRow(targetRowNum);
        if (targetRow == null) {
            targetRow = sheet.createRow(targetRowNum);
        }
        short srcHeight = rowData != null ? (short) rowData.getHeight() : sheet.getDefaultRowHeight();
        targetRow.setHeight(srcHeight);
    }
    
    /**
     * @return xls = null, xlsx = XSSFWorkbook, xlsx with streaming = the inner XSSFWorkbook instance
     */
    public XSSFWorkbook getXSSFWorkbook() {
        if (workbook instanceof SXSSFWorkbook) {
            return ((SXSSFWorkbook) workbook).getXSSFWorkbook();
        }
        if (workbook instanceof XSSFWorkbook) {
            return (XSSFWorkbook) workbook;
        }
        return null;
    }
    
    @Override
    public void adjustTableSize(CellRef ref, Size size) {
        XSSFWorkbook xwb = getXSSFWorkbook();
        if (size.getHeight() > 0 && xwb != null) {
            XSSFSheet sheet = xwb.getSheet(ref.getSheetName());
            if (sheet == null) {
                logger.error("Can not access sheet '{}'", ref.getSheetName());
            } else {
                for (XSSFTable table : sheet.getTables()) {
                    AreaRef areaRef = new AreaRef(table.getSheetName() + "!" + table.getCTTable().getRef());
                    if (areaRef.contains(ref)) {
                        // Make table higher
                        areaRef.getLastCellRef().setRow(ref.getRow() + size.getHeight() - 1);
                        table.getCTTable().setRef(
                                areaRef.getFirstCellRef().toString(true) + ":" + areaRef.getLastCellRef().toString(true));
                    }
                }
            }
        }
    }

    @Override
    public void mergeCells(CellRef cellRef, int rows, int cols) {
        Sheet sheet = getWorkbook().getSheet(cellRef.getSheetName());
        CellRangeAddress region = new CellRangeAddress(
                cellRef.getRow(),
                cellRef.getRow() + rows - 1,
                cellRef.getCol(),
                cellRef.getCol() + cols - 1);
        sheet.addMergedRegion(region);

        try {
            cellStyle = getCellStyle(cellRef);
        } catch (Exception ignore) {
        }
        for (int i = region.getFirstRow(); i <= region.getLastRow(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                row = sheet.createRow(i);
            }
            for (int j = region.getFirstColumn(); j <= region.getLastColumn(); j++) {
                Cell cell = row.getCell(j);
                if (cell == null) {
                    cell = row.createCell(j);
                }
                if (cellStyle == null) {
                    cell.getCellStyle().setAlignment(HorizontalAlignment.CENTER);
                    cell.getCellStyle().setVerticalAlignment(VerticalAlignment.CENTER);
                } else {
                    cell.setCellStyle(cellStyle);
                }
            }
        }
    }
}
