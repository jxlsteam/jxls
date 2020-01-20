package org.jxls.command;

@Deprecated
public class TestDepartment {
    private final String name;
    private final String key;
    
    public TestDepartment(String name, String key) {
        this.name = name;
        this.key = key;
    }

    public String getName() {
        return name;
    }

    // used by Issue133Test, issue133_template.xlsx
    public String getKey() {
        return key;
    }    
}
