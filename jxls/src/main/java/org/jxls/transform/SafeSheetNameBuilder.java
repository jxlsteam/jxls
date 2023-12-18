package org.jxls.transform;

import org.jxls.logging.JxlsLogger;

public interface SafeSheetNameBuilder {
    
    String CONTEXT_VAR_NAME = "SafeSheetNameBuilder";
    
    /**
     * @param givenSheetName -
     * @param index sheet index (starts with 0) can be used for the case that givenSheetName is null
     * @param logger -
     * @return safe version of givenSheetName
     */
    String createSafeSheetName(String givenSheetName, int index, JxlsLogger logger);
}
