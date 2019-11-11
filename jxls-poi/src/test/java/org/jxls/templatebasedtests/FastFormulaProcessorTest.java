package org.jxls.templatebasedtests;

import static org.mockito.ArgumentMatchers.any;
import static org.mockito.Mockito.spy;
import static org.mockito.Mockito.verify;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.junit.Assert;
import org.junit.Rule;
import org.junit.Test;
import org.jxls.EnglishTestRule;
import org.jxls.JxlsTester;
import org.jxls.common.CellData;
import org.jxls.common.CellRef;
import org.jxls.formula.FastFormulaProcessor;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.PoiTransformer;
import org.mockito.ArgumentCaptor;

/**
 * @author Michał Kępkowski
 */
public class FastFormulaProcessorTest {
    /** Makes the testcase work in a German environment where "IF" is called "WENN" in Excel. */
    @Rule
    public EnglishTestRule english = new EnglishTestRule();
    
    @Test
    public void testIfFormula() throws IOException {
        // Prepare
        Transformer transformer;
        try (InputStream is = JxlsTester.openInputStream(getClass(), false)) {
            try (OutputStream os = new ByteArrayOutputStream()) {
                transformer = spy(PoiTransformer.createTransformer(is, os)); // uses Mockito
            }
        }

        final String sheetName = "Arkusz1";
        transformer.getTargetCellRef(new CellRef(sheetName, 20, 4)).add(new CellRef(sheetName, 200, 30));
        transformer.getTargetCellRef(new CellRef(sheetName, 21, 4)).add(new CellRef(sheetName, 210, 30));
        List<CellData> cellDataList = new ArrayList<>(transformer.getFormulaCells());
        getIfFormula(cellDataList).addTargetPos(new CellRef(sheetName, 12, 12));

        // Test
        new FastFormulaProcessor().processAreaFormulas(transformer, null);

        // Verify
        ArgumentCaptor<String> firstFooCaptor = ArgumentCaptor.forClass(String.class);
        verify(transformer).setFormula(any(CellRef.class), firstFooCaptor.capture());
        String expected = "IF(AE201=0,\"\",AE211/AE201)";
        Assert.assertEquals(expected, firstFooCaptor.getAllValues().get(0));
    }

    private CellData getIfFormula(List<CellData> cellData) {
        for (CellData cell : cellData) {
            if (cell.getFormula().startsWith("IF")) {
                return cell;
            }
        }
        return null;
    }
}