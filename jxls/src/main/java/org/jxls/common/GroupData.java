package org.jxls.common;

import java.util.Collection;

/**
 * Used in collection grouping to hold group data
 * @author Leonid Vysochyn
 */
public class GroupData {
    private Object item;
    private Collection<Object> items;

    public GroupData(Object groupItem, Collection<Object> groupItems) {
        this.item = groupItem;
        this.items = groupItems;
    }

    public Object getItem() {
        return item;
    }

    public void setItem(Object item) {
        this.item = item;
    }

    public Collection<Object> getItems() {
        return items;
    }

    public void setItems(Collection<Object> items) {
        this.items = items;
    }
}
