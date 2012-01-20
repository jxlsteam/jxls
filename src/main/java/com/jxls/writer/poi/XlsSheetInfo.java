package com.jxls.writer.poi;

import com.jxls.writer.Pos;

import java.util.Map;
import java.util.HashMap;

/**
 * @author Leonid Vysochyn
 *         Date: 21.04.2009
 */
public class XlsSheetInfo {
    Map<Pos, XlsCellInfo> cells = new HashMap<Pos, XlsCellInfo>();

    public void add(Pos pos, XlsCellInfo cellInfo){
        cells.put( pos, cellInfo );
    }

    public XlsCellInfo getPos(Pos pos){
        return cells.get( pos );
    }

}
