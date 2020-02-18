package com.chenshuidesenlin.poiplus.core.xwpf.entity.chart;

import com.chenshuidesenlin.poiplus.core.xwpf.entity.Series;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

public abstract class AbstractChart implements Chart {

    private String title;

    private List<String> categories;

    private List<Series> series;

    @Override
    public String getTitle() {
        return title;
    }

    @Override
    public void setTitle(String title) {
        this.title = title;
    }

    @Override
    public List<String> getCategories() {
        return categories;
    }

    @Override
    public void setCategories(String... categoryArray) {
        List<String> categories = new ArrayList<>();
        Collections.addAll(categories, categoryArray);
        this.categories = categories;
    }

    @Override
    public void setCategories(List<String> categories) {
        this.categories = categories;
    }

    @Override
    public List<Series> getSeries() {
        return series;
    }

    @Override
    public void setSeries(List<Series> series) {
        this.series = series;
    }

    @Override
    public void addSeries(String name, Double... values) {
        if (series == null) {
            series = new ArrayList<>();
        }
        Series newSeries = new Series();
        newSeries.setName(name);
        newSeries.setValues(values);
        series.add(newSeries);
    }

    @Override
    public void addSeries(String name, List<Double> values) {
        if (series == null) {
            series = new ArrayList<>();
        }
        Series newSeries = new Series();
        newSeries.setName(name);
        newSeries.setValues(values);
        series.add(newSeries);
    }

}
