package org.jxls.expression;

/**
 * @author Leonid Vysochyn
 */
public class Dummy {
    private String strValue;
    private int intValue;

    public Dummy(String strValue, int intValue) {
        this.strValue = strValue;
        this.intValue = intValue;
    }

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

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        Dummy dummy = (Dummy) o;
        if (intValue != dummy.intValue) return false;
        if (strValue != null ? !strValue.equals(dummy.strValue) : dummy.strValue != null) return false;
        return true;
    }

    @Override
    public int hashCode() {
        int result = strValue != null ? strValue.hashCode() : 0;
        result = 31 * result + intValue;
        return result;
    }
}
