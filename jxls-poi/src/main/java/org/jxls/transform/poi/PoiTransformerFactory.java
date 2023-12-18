package org.jxls.transform.poi;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashSet;
import java.util.Set;

import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.jxls.builder.JxlsStreaming;
import org.jxls.logging.JxlsLogger;
import org.jxls.transform.JxlsTransformerFactory;
import org.jxls.transform.Transformer;
import org.jxls.util.CannotOpenWorkbookException;

public class PoiTransformerFactory implements JxlsTransformerFactory {

    @Override
    public Transformer create(InputStream template, OutputStream outputStream, JxlsStreaming streaming, JxlsLogger logger) {
        Workbook workbook = openWorkbook(template);
        PoiTransformer transformer = createTransformer(workbook, streaming);
        transformer.setOutputStream(outputStream);
        return transformer;
    }

    protected Workbook openWorkbook(InputStream template) {
        try {
            return WorkbookFactory.create(template);
        } catch (Exception e) {
            throw new CannotOpenWorkbookException(e);
        }
    }

    protected PoiTransformer createTransformer(Workbook workbook, JxlsStreaming streaming) {
        if (streaming.isAutoDetect()) {
            return new SelectSheetsForStreamingPoiTransformer(workbook, getAllSheetsInWhichStreamingIsConfigured(workbook));
        } else if (streaming.getSheetNames() != null) {
            return new SelectSheetsForStreamingPoiTransformer(workbook, streaming.getSheetNames());
        } else if (streaming.isStreaming()) {
            // Don't use PoiTransformer here because SelectSheetsForStreamingPoiTransformer is better.
            return new SelectSheetsForStreamingPoiTransformer(workbook, true);
        } else {
            return new PoiTransformer(workbook, false);
        }
    }

    public static Set<String> getAllSheetsInWhichStreamingIsConfigured(Workbook workbook) {
        Set<String> sheetNames = new HashSet<>();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            if (isStreamingEnabled(sheet)) {
                sheetNames.add(sheet.getSheetName());
            }
        }
        return sheetNames;
    }

    private static boolean isStreamingEnabled(Sheet sheet) {
        for (Comment comment : sheet.getCellComments().values()) {
            String text = comment.getString().getString();
            if (text.contains("sheetStreaming=\"true\"")) {
                return true;
            }
        }
        return false;
    }
}
