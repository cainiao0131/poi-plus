package com.chenshuidesenlin.poiplus.core.xwpf.entity;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

public class Series {

    private String name;

    private List<Double> values;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<Double> getValues() {
        return values;
    }

    public void setValues(List<Double> values) {
        this.values = values;
    }

    public void setValues(Double[] values) {
        List<Double> newValues = new ArrayList<>();
        Collections.addAll(newValues, values);
        this.values = newValues;
    }

}
