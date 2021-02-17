package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.io.IOException;

import org.apache.poi.ss.usermodel.RichTextString;
import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;

/**
 * Issue with formatting of parts of text
 */
public class IssueB107Test {

    @Test
    public void test() throws IOException {
        // Test
        JxlsTester tester = JxlsTester.xls(getClass());
        tester.processTemplate(new Context());
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            RichTextString r = w.getCellValueAsRichString(2, 1);
            assertEquals("ABC", r.toString()); // in template and output file: "A" is black and "BC" is red
            assertEquals("Rich text string must consist of 2 parts.", 2, r.numFormattingRuns());
        }
    }
}
