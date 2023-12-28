package org.jxls.template;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.command.GridCommand;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.common.JxlsException;
import org.jxls.expression.ExpressionEvaluatorFactoryJexlImpl;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.PoiTransformer;
import org.jxls.util.CannotOpenWorkbookException;

/**
 * Helper class for grid export
 */
public class SimpleExporter {
    public static final String GRID_TEMPLATE_XLS = "grid_template.xls";

    private byte[] templateBytes;

    public SimpleExporter() {
    }

    public void registerGridTemplate(InputStream inputStream) throws IOException {
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        byte[] data = new byte[4096];
        int count;
        while ((count = inputStream.read(data)) != -1) {
            os.write(data, 0, count);
        }
        templateBytes = os.toByteArray();
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
    	InputStream is = new ByteArrayInputStream(templateBytes);
        Transformer transformer = createTransformer(is, outputStream);
        transformer.getTransformationConfig().setExpressionEvaluatorFactory(new ExpressionEvaluatorFactoryJexlImpl());
        AreaBuilder areaBuilder = new XlsCommentAreaBuilder();
        List<Area> xlsAreaList = areaBuilder.build(transformer, true);
        Area xlsArea = xlsAreaList.get(0);
        Context context = new Context();
        context.putVar("headers", headers);
        context.putVar("data", dataObjects);
        GridCommand gridCommand = (GridCommand) xlsArea.getCommandDataList().get(0).getCommand();
        gridCommand.setProps(objectProps);
        xlsArea.applyAt(new CellRef("Sheet1!A1"), context);
        try {
            transformer.write();
        } catch (IOException e) {
            throw new JxlsException("Failed to write to output stream", e);
        }
    }

    private PoiTransformer createTransformer(InputStream is, OutputStream os) {
        try {
            Workbook workbook = WorkbookFactory.create(is);
            PoiTransformer transformer = new PoiTransformer(workbook, false);
            transformer.setOutputStream(os);
            return transformer;
        } catch (EncryptedDocumentException | IOException e) {
            throw new CannotOpenWorkbookException(e);
        }
    }
}
