package org.jxls;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Assert;
import org.jxls.builder.JxlsTemplateFillerBuilder;
import org.jxls.transform.poi.PoiTransformer;
import org.jxls.util.CannotOpenWorkbookException;

public class Jxls3Tester {
    private final Class<?> testclass;
    private final String excelTemplateFilename;
    private final File out;

    public static Jxls3Tester xlsx(Class<?> testclass) {
        return new Jxls3Tester(testclass, testclass.getSimpleName() + ".xlsx");
    }

    public static Jxls3Tester xlsx(Class<?> testclass, String method) {
        return new Jxls3Tester(testclass, testclass.getSimpleName() + "_" + method + ".xlsx");
    }

    public Jxls3Tester(Class<?> testclass, String excelTemplateFilename) {
        this.testclass = testclass;
        this.excelTemplateFilename = excelTemplateFilename;
        int o = excelTemplateFilename.lastIndexOf(".");
        File folder = new File("target");
        folder.mkdir();
        out = new File(folder, excelTemplateFilename.substring(0, o) + "_output" + excelTemplateFilename.substring(o));
    }
    
    public void test(Map<String, Object> data, JxlsTemplateFillerBuilder<?> builder) {
        InputStream template = testclass.getResourceAsStream(excelTemplateFilename);
        Assert.assertNotNull("Template not found: " + excelTemplateFilename, template);
        builder.withTemplate(template).buildAndFill(data, out);
    }
    
    public TestWorkbook getWorkbook() {
        return new TestWorkbook(out);
    }
    
    public static PoiTransformer createTransformer(InputStream is, OutputStream os) {
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
