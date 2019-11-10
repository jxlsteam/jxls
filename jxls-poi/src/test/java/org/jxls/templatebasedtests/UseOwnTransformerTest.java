package org.jxls.templatebasedtests;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Assert;
import org.junit.Test;
import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.transformer.OwnTransformer;
import org.jxls.util.JxlsHelper;

/**
 * This testcase ensures that PoiTransformer is replaceable or extendable.
 */
public class UseOwnTransformerTest {

    @Test
    public void testReplacementOfPoiTransformer() throws EncryptedDocumentException, IOException {
        InputStream in = UseOwnTransformerTest.class.getResourceAsStream("UseOwnTransformerTest.xlsx"); // simple XLSX file with 1 jx:area and 1 jx:each.
        try {
            Workbook workbook = WorkbookFactory.create(in);
            try {
                Transformer transformer = new OwnTransformer(workbook); // use other implementation of PoiTransformer
                ByteArrayOutputStream out = new ByteArrayOutputStream();
                try {
                    ((OwnTransformer) transformer).setOutputStream(out);
                    JxlsHelper.getInstance().processTemplate(new Context(), transformer);
                } finally {
                    out.close();
                }
                
                Assert.assertTrue("Must be OwnTransformer", transformer instanceof OwnTransformer);
                Assert.assertTrue("clearCell() must be called", ((OwnTransformer) transformer).isClearCellCalled());
                Assert.assertTrue("transform() must be called", ((OwnTransformer) transformer).isTransformCalled());
            } finally {
                workbook.close();
            }
        } finally {
            in.close();
        }
    }
}
