package org.jxls.command;

import javax.xml.namespace.QName;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.jxls.area.Area;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.JxlsException;
import org.jxls.common.Size;
import org.jxls.transform.poi.PoiTransformer;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.impl.CTRowImpl;

/**
 * <p>
 * Automatically adjust rows height to fit content.
 * </p>
 * <p>
 * This is obtained setting row height to -1.
 * </p>
 *
 * @see https://stackoverflow.com/a/48789397/5116356
 * @since 3.1.0
 */
public class AutoRowHeightCommand extends AbstractCommand {
    public static final String COMMAND_NAME = "autoRowHeight";

    private Area area;

    @Override
    public String getName() {
        return COMMAND_NAME;
    }

    @Override
    public Command addArea(Area area) {
        if (area == null) {
            return this;
        }
        if (!areaList.isEmpty()) {
            throw new JxlsException("You can add only a single area to 'autoRowHeight' command");
        }
        this.area = area;
        return super.addArea(area);
    }

    @Override
    public Size applyAt(CellRef cellRef, Context context) {
        Size size = area.applyAt(cellRef, context);
        PoiTransformer transformer = (PoiTransformer) area.getTransformer();
        Row row = transformer.getWorkbook().getSheet(cellRef.getSheetName()).getRow(cellRef.getRow());
        removeDyDescentAttr(row);
        row.setHeight((short) -1);
        return size;
    }

    /**
     * Workaround for dyDescent attribute
     * 
     * @see https://stackoverflow.com/a/53782199/5116356
     */
    private void removeDyDescentAttr(Row row) {
        try {
            XSSFRow xssfRow = (XSSFRow) row;
            CTRowImpl ctRow = (CTRowImpl) xssfRow.getCTRow();
            QName dyDescent = new QName("http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            if (ctRow.get_store().find_attribute_user(dyDescent) != null) {
                ctRow.get_store().remove_attribute(dyDescent);
            }
        } catch (ClassCastException e) {
            // do nothing, probably not an XSSF sheet
        }
    }
}
