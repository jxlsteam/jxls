package org.jxls.entity;

import java.util.List;

public class Item {
    private long id;
    private String name;
    private List<CS> cs;

    public Item(long id, String name) {
        this.id = id;
        this.name = name;
    }

    public Item(long id, String name, List<CS> cs) {
        this.id = id;
        this.name = name;
        this.cs = cs;
    }

    public long getId() {
        return id;
    }

    public void setId(long id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<CS> getCs() {
        return cs;
    }

    public void setCs(List<CS> cs) {
        this.cs = cs;
    }
}
