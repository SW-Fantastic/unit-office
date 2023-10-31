package org.swdc.offices.generator;

import java.util.*;
import java.util.stream.Collectors;

public class PipedGenerationContext {

    private Map<Class, List<Object>> objects = new HashMap<>();

    private List<Object> resolved = new ArrayList<>();

    public PipedGenerationContext(List<? extends Object> objects) {
        this.objects =  objects.stream().filter(Objects::nonNull)
                .collect(Collectors.groupingBy(Object::getClass));
    }

    public <E> List<E> getGrouped(Class<E> type) {
        if (objects.containsKey(type)) {
            return (List<E>) objects.get(type);
        }
        for (Class curType: objects.keySet()) {
            if (type.isAssignableFrom(curType)) {
                return (List<E>) objects.get(curType);
            }
        }
        return Collections.emptyList();
    }

    public <E> E singleton(Class<E> type) {
        List<E> items = getGrouped(type);
        if (items.size() == 1) {
            return items.get(0);
        } else if (items.size() > 1) {
            throw new RuntimeException("there are more than one object with type:" + type.getName());
        }
        return null;
    }

    public void resolved(Object object) {
        for (List<Object> vals : objects.values()) {
            if (vals.contains(object)) {
                vals.remove(object);
                resolved.add(object);
                break;
            }
        }
    }

}
