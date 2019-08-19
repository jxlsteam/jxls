package org.jxls.transform;

public interface SafeSheetNameBuilder {
    
    String CONTEXT_VAR_NAME = "SafeSheetNameBuilder";
    
    /**
     * @param givenSheetName
     * @return safe version of givenSheetName
     */
    String createSafeSheetName(String givenSheetName);
}
