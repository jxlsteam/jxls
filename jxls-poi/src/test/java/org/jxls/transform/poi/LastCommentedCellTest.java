package org.jxls.transform.poi;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Assert;
import org.junit.Test;
import org.jxls.common.CellData;

public class LastCommentedCellTest {

    @Test
    public void lastCommentedCell() throws EncryptedDocumentException, IOException {
        // Prepare
        try (InputStream is = getClass().getResourceAsStream(getClass().getSimpleName() + ".xlsx")) {
            Workbook workbook = WorkbookFactory.create(is);
            PoiTransformer transformer = new PoiTransformer(workbook, false);
            
            // Test
            List<CellData> cc = transformer.getCommentedCells();
            
            // Verify
            Assert.assertTrue(cc.stream().anyMatch(i -> i.getCellComment().contains("comment-1")));
            Assert.assertTrue(cc.stream().anyMatch(i -> i.getCellComment().contains("comment in AZ")));
        }
    }
}
