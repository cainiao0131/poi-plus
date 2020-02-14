package com.cainiao.poiplus.core.xwpf.dto.chart;

import com.cainiao.poiplus.core.xwpf.dto.Series;

import java.util.List;

public interface Chart {

    String getTitle();

    Chart setTitle(String title);

    String[] getCategories();

    Chart setCategories(String... categories);

    List<Series> getSeries();

    void setSeries(List<Series> series);

    Chart addSeries(String name, Double... values);

}
