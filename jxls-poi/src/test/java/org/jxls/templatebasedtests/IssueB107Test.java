package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.io.IOException;

import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.ContextImpl;

/**
 * Issue with formatting of parts of text
 */
public class IssueB107Test {

    @Test
    public void test() throws IOException {
        // Test
        JxlsTester tester = JxlsTester.xls(getClass());
        tester.processTemplate(new ContextImpl());
        
        // Verify
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            assertEquals("ABC", w.getCellValueAsRichString(2, 1)); // in template and output file: "A" is black and "BC" is red
            assertEquals("Rich text string must consist of 2 parts.", 2, w.getCellValueAsRichStringNumFormattingRuns(2, 1));
        }
    }
}
