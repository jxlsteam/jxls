package org.jxls.expression;

import org.apache.commons.jexl3.introspection.JexlPermissions;

public class JxlsJexlPermissions {
    public static final JxlsJexlPermissions UNRESTRICTED = new JxlsJexlPermissions(0);
    public static final JxlsJexlPermissions   RESTRICTED = new JxlsJexlPermissions(1);
    private final int id;
    private final JexlPermissions jexlPermissions;
    
    public JxlsJexlPermissions(String ...src) {
        String sum = ""; // no StringBuilder!
        for (String line : src) {
            sum += "" + line.hashCode();
        }
        id = sum.hashCode();
        jexlPermissions = JexlPermissions.parse(src);
    }
    
    private JxlsJexlPermissions(int type) {
        jexlPermissions = type == 0 ? JexlPermissions.UNRESTRICTED : JexlPermissions.RESTRICTED;
        id = jexlPermissions.hashCode();
    }
    
    @Override
    public int hashCode() {
        return id;
    }
    
    @Override
    public boolean equals(Object o) {
        return o instanceof JxlsJexlPermissions && ((JxlsJexlPermissions) o).id == id;
    }
    
    JexlPermissions getJexlPermissions() {
        return jexlPermissions;
    }
}
