package org.jxls.template;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.jxls.area.Area;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.command.GridCommand;
import org.jxls.command.ImageCommand;
import org.jxls.common.JxlsException;
import org.jxls.transform.poi.JxlsPoiTemplateFillerBuilder;

/**
 * Helper class for grid export
 */
public class SimpleExporter {
    public static final String GRID_TEMPLATE_XLS = "grid_template.xlsx";
    private byte[] templateBytes;

    public void registerGridTemplate(InputStream inputStream) throws IOException {
        templateBytes = ImageCommand.toByteArray(inputStream);
    }

    public void gridExport(Iterable<?> headers, Iterable<?> dataObjects, String objectProps, OutputStream outputStream) {
        if (templateBytes == null) {
            InputStream is = SimpleExporter.class.getResourceAsStream(GRID_TEMPLATE_XLS);
            try {
                registerGridTemplate(is);
            } catch (IOException e) {
                throw new JxlsException("Failed to read default template file " + GRID_TEMPLATE_XLS, e);
            }
    	}    	
        Map<String, Object> data = new HashMap<>();
        data.put("headers", headers);
        data.put("data", dataObjects);
    	JxlsPoiTemplateFillerBuilder.newInstance().withTemplate(new ByteArrayInputStream(templateBytes))
    	    .withAreaBuilder((transformer, ctc) -> {
    	        List<Area> areas = new XlsCommentAreaBuilder().build(transformer, ctc);
    	        GridCommand gridCommand = (GridCommand) areas.get(0).getCommandDataList().get(0).getCommand();
    	        gridCommand.setProps(objectProps);
    	        return areas;
    	    })
    	    .buildAndFill(data, () -> outputStream);
    }
}
