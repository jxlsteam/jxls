package com.jxls.writer.expression;

/**
 * @author Leonid Vysochyn
 *         Date: 1/30/12 1:26 PM
 */
public class Dummy {
    String strValue;
    int intValue;

    public Dummy(int intValue) {
        this.intValue = intValue;
    }

    public Dummy(String strValue) {
        this.strValue = strValue;
    }

    public int getIntValue() {
        return intValue;
    }

    public void setIntValue(int intValue) {
        this.intValue = intValue;
    }

    public String getStrValue() {
        return strValue;
    }

    public void setStrValue(String strValue) {
        this.strValue = strValue;
    }
}
