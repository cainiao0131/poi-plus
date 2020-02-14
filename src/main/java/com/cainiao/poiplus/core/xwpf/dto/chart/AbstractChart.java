package com.cainiao.poiplus.core.xwpf.dto.chart;

import com.cainiao.poiplus.core.xwpf.dto.Series;

import java.util.ArrayList;
import java.util.List;

public abstract class AbstractChart implements Chart {

    private String title;

    private String[] categories;

    private List<Series> series;

    @Override
    public String getTitle() {
        return title;
    }

    @Override
    public AbstractChart setTitle(String title) {
        this.title = title;
        return this;
    }

    @Override
    public String[] getCategories() {
        return categories;
    }

    @Override
    public AbstractChart setCategories(String... categories) {
        this.categories = categories;
        return this;
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
    public AbstractChart addSeries(String name, Double... values) {
        if (series == null) {
            series = new ArrayList<>();
        }
        Series newSeries = new Series();
        newSeries.setName(name);
        newSeries.setValues(values);
        series.add(newSeries);
        return this;
    }

}
