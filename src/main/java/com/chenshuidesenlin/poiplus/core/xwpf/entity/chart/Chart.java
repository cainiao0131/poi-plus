package com.chenshuidesenlin.poiplus.core.xwpf.entity.chart;

import com.chenshuidesenlin.poiplus.core.xwpf.entity.Series;
import com.chenshuidesenlin.poiplus.core.xwpf.entity.WordContent;

import java.util.List;

public interface Chart extends WordContent {

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
