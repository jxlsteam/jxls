package com.jxls.writer.builder.xls;

import com.jxls.writer.area.Area;
import com.jxls.writer.builder.AreaBuilder;
import com.jxls.writer.common.CellData;
import com.jxls.writer.transform.Transformer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @author Leonid Vysochyn
 */
public class XlsCommentAreaBuilder implements AreaBuilder {
    Transformer transformer;

    public XlsCommentAreaBuilder(Transformer transformer) {
        this.transformer = transformer;
    }

    public List<Area> build() {
        List<Area> areas = new ArrayList<Area>();
        List<CellData> commentedCells = transformer.getCommentedCells();
        for (CellData cellData : commentedCells) {
            String comment = cellData.getCellComment();
        }
        return areas;
    }

}
