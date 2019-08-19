package org.jxls.transform.poi;

import org.apache.poi.ss.util.WorkbookUtil;
import org.jxls.transform.SafeSheetNameBuilder;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Ensures valid Excel sheet names.
 * Use: Put a class instance with name SafeSheetNameBuilder.CONTEXT_VAR_NAME into the Context.
 */
public class PoiSafeSheetNameBuilder implements SafeSheetNameBuilder {
    private static Logger logger = LoggerFactory.getLogger(PoiSafeSheetNameBuilder.class);

    @Override
    public String createSafeSheetName(String givenSheetName) {
        String sheetName = WorkbookUtil.createSafeSheetName(givenSheetName);
        if (!givenSheetName.equals(sheetName)) {
            logger.info("Change invalid sheet name {} to {}", givenSheetName, sheetName);
        }
        return sheetName;
    }
}
