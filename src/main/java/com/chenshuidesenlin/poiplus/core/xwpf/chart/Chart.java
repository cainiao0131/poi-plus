package com.chenshuidesenlin.poiplus.core.xwpf.chart;

import com.chenshuidesenlin.poiplus.core.xwpf.entity.Series;

import java.util.List;

public interface Chart {

    String getTitle();

    void setTitle(String title);

    List<String> getCategories();

    void setCategories(String... categories);

    void setCategories(List<String> categories);

    List<Series> getSeries();

    void setSeries(List<Series> series);

    void addSeries(String name, Double... values);

    void addSeries(String name, List<Double> values);

}
