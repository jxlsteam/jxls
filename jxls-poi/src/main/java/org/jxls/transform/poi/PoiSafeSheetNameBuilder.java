package org.jxls.transform.poi;

import java.util.HashSet;
import java.util.Set;

import org.apache.poi.ss.util.WorkbookUtil;
import org.jxls.transform.SafeSheetNameBuilder;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Ensures valid and unique Excel sheet names.
 * Use: For every Excel file put a new class instance with name SafeSheetNameBuilder.CONTEXT_VAR_NAME into the Context.
 */
public class PoiSafeSheetNameBuilder implements SafeSheetNameBuilder {
    private static final Logger logger = LoggerFactory.getLogger(PoiSafeSheetNameBuilder.class);
    private final Set<String> usedSheetNames = new HashSet<>();
    
    @Override
    public String createSafeSheetName(final String givenSheetName, int index) {
        String sheetName = WorkbookUtil.createSafeSheetName(givenSheetName);
        int serialNumber = getFirstSerialNumber();
        String newName = sheetName;
        while (usedSheetNames.contains(newName)) { // until unique
            int len = sheetName.length();
            String nameWithNumber;
            do {
                nameWithNumber = addSerialNumber(sheetName.substring(0, len--), serialNumber);
                newName = WorkbookUtil.createSafeSheetName(nameWithNumber);
            } while (!newName.equals(nameWithNumber)); // while createSafeSheetName() changes the name
            serialNumber++;
        }
        if (!givenSheetName.equals(newName)) {
            logger.info("Change invalid sheet name {} to {}", givenSheetName, sheetName);
        }
        usedSheetNames.add(newName);
        return newName;
    }

    protected int getFirstSerialNumber() {
        return 1;
    }

    protected String addSerialNumber(String text, int serialNumber) {
        return text + "(" + serialNumber + ")";
    }
}
