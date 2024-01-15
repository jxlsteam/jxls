package org.jxls.common;

import java.util.Collection;
import java.util.Map;
import java.util.Set;

/**
 * Map implementation that delegates to another map for use as decorator pattern.
 */
public class MapDelegator<K, V> implements Map<K, V> {
    private final Map<K, V> map;
    
    public MapDelegator(Map<K, V> map) {
        if (map == null) {
            throw new IllegalArgumentException("map must not be null");
        }
        this.map = map;
    }

    // not used by Jxls
    @Override
    public int size() {
        return map.size();
    }

    // not used by Jxls
    @Override
    public boolean isEmpty() {
        return map.isEmpty();
    }

    @Override
    public boolean containsKey(Object key) {
        return map.containsKey(key);
    }

    // not used by Jxls
    @Override
    public boolean containsValue(Object value) {
        return map.containsValue(value);
    }

    @Override
    public V get(Object key) {
        return map.get(key);
    }

    @Override
    public V put(K key, V value) {
        return map.put(key, value);
    }

    @Override
    public V remove(Object key) {
        return map.remove(key);
    }

    // not used by Jxls
    @Override
    public void putAll(Map<? extends K, ? extends V> m) {
        map.putAll(m);
    }

    // not used by Jxls
    @Override
    public void clear() {
        map.clear();
    }

    // not used by Jxls
    @Override
    public Set<K> keySet() {
        return map.keySet();
    }

    // not used by Jxls
    @Override
    public Collection<V> values() {
        return map.values();
    }

    // used by JSR223
    @Override
    public Set<Entry<K, V>> entrySet() {
        return map.entrySet();
    }
    
    @Override
    public String toString() {
        return map.toString();
    }

    @Override
    public boolean equals(Object object) {
        return object == this || map.equals(object);
    }

    @Override
    public int hashCode() {
        return map.hashCode();
    }
}
