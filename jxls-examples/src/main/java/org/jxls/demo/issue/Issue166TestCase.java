package org.jxls.demo.issue;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.formula.StandardFormulaProcessor;
import org.jxls.transform.Transformer;
import org.jxls.util.JxlsHelper;

/**
 * Wrong average on 2nd sheet
 */
public class Issue166TestCase {
    
    // --------- SETTINGS ---------

    final static String INPUT_FILE_PATH = "issue166_Template.xlsx";
    final static String OUTPUT_FILE_PATH = "target/Issue166_Output.xlsx";

    // --------- -------- ---------

    public static void main(String[] args) throws IOException {
    	
    	// define result set
        List<Map<String, Object>> rs = new ArrayList<>();
        for (int i = 0; i < 5; i++) {
            Map<String, Object> map = new HashMap<String, Object>();
            map.put("count", i);
            rs.add(map);
        }
    	
        try(InputStream is = Issue166TestCase.class.getResourceAsStream(INPUT_FILE_PATH)) {
            try (OutputStream os = new FileOutputStream(OUTPUT_FILE_PATH)) {
            	JxlsHelper jxlsHelper = JxlsHelper.getInstance();
                Context context = new Context();
                context.putVar("rs0", rs);
                
                Transformer transformer = jxlsHelper.createTransformer(is, os);
        		AreaBuilder areaBuilder = new XlsCommentAreaBuilder();
        		areaBuilder.setTransformer(transformer);
                List<Area> xlsAreaList = areaBuilder.build();
                for (Area xlsArea : xlsAreaList) {
                    xlsArea.applyAt(new CellRef(xlsArea.getStartCellRef().getCellName()), context);
                    // Make sure the StandardFormulaProcessor is used. This will make sure formula references on multi-sheet workbooks are correct.
                    xlsArea.setFormulaProcessor(new StandardFormulaProcessor());
                    xlsArea.processFormulas();            
                }
        		jxlsHelper.processTemplate(context, transformer);
            }
        }
    }
}
