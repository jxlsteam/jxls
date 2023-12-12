package org.jxls.builder;

import java.util.Arrays;
import java.util.HashSet;
import java.util.Set;

/**
 * You have 4 options for streaming: STREAMING_OFF, STREAMING_ON, AUTO_DETECT or streamingWithGivenSheets(...)
 */
public class JxlsStreaming {
    private final boolean streaming;
    private final Set<String> sheetNames;
    private final boolean autoDetect;

    public static final JxlsStreaming STREAMING_OFF = new JxlsStreaming(false, null, false);
    /** Use standard streaming for all sheets */
    public static final JxlsStreaming STREAMING_ON = new JxlsStreaming(true, null, false);
    /** All sheets where a note contain exactly sheetStreaming="true" will be used for streaming. */
    public static final JxlsStreaming AUTO_DETECT = new JxlsStreaming(true, null, true);

    public static JxlsStreaming streamingWithGivenSheets(Set<String> sheetNames) {
        return new JxlsStreaming(true, sheetNames, false);
    }

    public static JxlsStreaming streamingWithGivenSheets(String ...sheetNames) {
        return new JxlsStreaming(true, new HashSet<>(Arrays.asList(sheetNames)), false);
    }

    private JxlsStreaming(boolean streaming, Set<String> sheetNames, boolean autoDetect) {
        this.streaming = streaming;
        this.sheetNames = sheetNames;
        this.autoDetect = autoDetect;
    }

    public boolean isStreaming() {
        return streaming;
    }

    public Set<String> getSheetNames() {
        return sheetNames;
    }

    public boolean isAutoDetect() {
        return autoDetect;
    }
}
