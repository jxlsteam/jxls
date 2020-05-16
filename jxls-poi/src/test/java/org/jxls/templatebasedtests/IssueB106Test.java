package org.jxls.templatebasedtests;

import static org.junit.Assert.assertEquals;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;
import org.jxls.JxlsTester;
import org.jxls.TestWorkbook;
import org.jxls.common.Context;

/**
 * Issue with row height and jx:area not in A1
 */
public class IssueB106Test {

    @Test
    public void test() throws IOException {
        // Test
        JxlsTester tester = JxlsTester.xls(getClass());
        tester.processTemplate(new Context());

        // Verify heights of rows 2 and 3
        try (TestWorkbook w = tester.getWorkbook()) {
            w.selectSheet(0);
            assertEquals("ABC", w.getCellValueAsString(3, 1));
            try (Workbook template = openTemplate(tester)) {
                for (int row = 2; row <= 3; row++) {
                    assertEquals("Wrong height of row " + row, getRowHeight(template, row), w.getRowHeight(row));
                }
            }
        }
    }

    private Workbook openTemplate(JxlsTester tester) throws IOException {
        return WorkbookFactory.create(getClass().getResourceAsStream(tester.getTemplateFilename()));
    }
    
    private short getRowHeight(Workbook workbook, int row) {
        return workbook.getSheetAt(0).getRow(row - 1).getHeight();
    }
}
