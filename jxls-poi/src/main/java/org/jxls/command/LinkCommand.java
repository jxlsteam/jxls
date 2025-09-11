package org.jxls.command;

import java.util.Objects;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.jxls.area.Area;
import org.jxls.common.AreaListener;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.Size;
import org.jxls.transform.poi.PoiTransformer;

/**
 * Hyperlink rendering
 * 
 * @author Wangtd
 */
public class LinkCommand extends AbstractCommand {
    public static final String COMMAND_NAME = "link";

    private Area area;
    private String href;
    private String type;
    private String label;

    @Override
    public String getName() {
        return COMMAND_NAME;
    }

    public void setHref(String href) {
        this.href = href;
    }

    public void setType(String type) {
        this.type = type;
    }

    public void setLabel(String label) {
        this.label = label;
    }

    @Override
    public Command addArea(Area area) {
        if (areaList.size() >= 1) {
            throw new IllegalArgumentException("You can only add 1 area to 'link' command!");
        }
        this.area = area;
        this.area.addAreaListener(new AreaListener() {
            @Override
            public void beforeTransformCell(CellRef srcCell, CellRef targetCell, Context context) {
                // no-op
            }

            @Override
            public void afterTransformCell(CellRef srcCell, CellRef targetCell, Context context) {
                // no-op
            }

            @Override
            public void beforeApplyAtCell(CellRef cellRef, Context context) {
                // no-op
            }

            @Override
            public void afterApplyAtCell(CellRef cellRef, Context context) {
                String linkHref = null;
                if (Objects.nonNull(href)) {
                    linkHref = (String) context.evaluate(href);
                }
                if (Objects.isNull(linkHref)) {
                    return;
                }
                String linkLabel = null;
                if (Objects.nonNull(label)) {
                    linkLabel = (String) context.evaluate(label);
                }
                HyperlinkType linkType = HyperlinkType.URL;
                if (Objects.nonNull(type)) {
                    int typeIndex = (Integer) context.evaluate(type);
                    linkType = HyperlinkType.values()[typeIndex];
                }

                PoiTransformer transformer = (PoiTransformer) getTransformer();
                Workbook workbook = transformer.getWorkbook();
                Sheet sheet = workbook.getSheet(cellRef.getSheetName());
                if (Objects.isNull(sheet)) {
                    sheet = workbook.getSheet(cellRef.getSheetName());
                }
                Row row = sheet.getRow(cellRef.getRow());
                if (Objects.isNull(row)) {
                    row = sheet.createRow(cellRef.getRow());
                }
                Cell cell = row.getCell(cellRef.getCol());
                if (Objects.isNull(cell)) {
                    cell = row.createCell(cellRef.getCol());
                }
                // Set value for cell using linkLabel
                if (Objects.nonNull(linkLabel)) {
                    cell.setCellValue(linkLabel);
                }
                // Create hyperlink by type
                CreationHelper helper = workbook.getCreationHelper();
                Hyperlink link = helper.createHyperlink(linkType);
                link.setAddress(linkHref);
                if (Objects.nonNull(linkLabel)) {
                    link.setLabel(linkLabel);
                }
                cell.setHyperlink(link);
                // Add hyperlink style for cell
                CellStyle newStyle = workbook.createCellStyle();
                CellStyle cellStyle = cell.getCellStyle();
                if (Objects.nonNull(cellStyle)) {
                    newStyle.cloneStyleFrom(cellStyle);
                }
                Font font = workbook.createFont();
                font.setUnderline(Font.U_SINGLE);
                font.setColor(IndexedColors.BLUE.getIndex());
                newStyle.setFont(font);
                cell.setCellStyle(newStyle);
            }
        });
        return super.addArea(area);
    }

    @Override
    public Size applyAt(CellRef cellRef, Context context) {
        if (Objects.isNull(area)) {
            throw new IllegalArgumentException("No area is defined for link command");
        }
        // Apply hyperlink cell context
        return area.applyAt(cellRef, context);
    }
}
