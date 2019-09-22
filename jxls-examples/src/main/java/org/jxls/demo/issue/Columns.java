package org.jxls.demo.issue;

import java.util.Collection;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

public class Columns {
    public Collection<String> keyOf(List<Map<String, Object>> row) {
        return row.get(0).keySet().stream().filter(k -> k.endsWith("_dynamic")).collect(Collectors.toList());
    }

    public Collection<Object> valueOf(Map<String, Object> row) {
        return row.entrySet().stream()
                .filter(entry -> entry.getKey() != null && entry.getKey().endsWith("_dynamic"))
                .map(Map.Entry::getValue)
                .collect(Collectors.toList());
    }
}
