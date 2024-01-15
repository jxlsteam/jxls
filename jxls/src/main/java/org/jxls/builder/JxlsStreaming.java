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
    /** option for streaming; see SXSSFWorkbook.DEFAULT_WINDOW_SIZE */
    private int rowAccessWindowSize = 100;
    private boolean compressTmpFiles = false;
    private boolean useSharedStringsTable = false;

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

    /**
     * Not often used streaming options
     * @param rowAccessWindowSize default 100, see SXSSFWorkbook.DEFAULT_WINDOW_SIZE
     * @param compressTmpFiles default false
     * @param useSharedStringsTable default false
     * @return new JxlsStreaming object
     */
    public JxlsStreaming withOptions(int rowAccessWindowSize, boolean compressTmpFiles, boolean useSharedStringsTable) {
        JxlsStreaming ret = new JxlsStreaming(streaming, sheetNames, autoDetect);
        ret.rowAccessWindowSize = rowAccessWindowSize;
        ret.compressTmpFiles = compressTmpFiles;
        ret.useSharedStringsTable = useSharedStringsTable;
        return ret;
    }

    public int getRowAccessWindowSize() {
        return rowAccessWindowSize;
    }

    public boolean isCompressTmpFiles() {
        return compressTmpFiles;
    }

    public boolean isUseSharedStringsTable() {
        return useSharedStringsTable;
    }
}
