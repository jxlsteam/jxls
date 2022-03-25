package org.jxls.util;

public abstract class BooleanLiteral {

}

class SmallT extends BooleanLiteral {

    static boolean isaBooleanT(String rawSheetName) {
        return "TRUE".equalsIgnoreCase(rawSheetName);
    }
}

class SmallF extends BooleanLiteral {
    static boolean isaBooleanF(String rawSheetName) {
        return "FALSE".equalsIgnoreCase(rawSheetName);
    }
}
