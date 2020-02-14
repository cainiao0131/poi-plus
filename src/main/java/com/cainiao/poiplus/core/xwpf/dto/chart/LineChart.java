package com.cainiao.poiplus.core.xwpf.dto.chart;

/**
 * 折线图
 */
public class LineChart extends AbstractChart {

    @Override
    public LineChart setTitle(String title) {
        super.setTitle(title);
        return this;
    }

    @Override
    public LineChart setCategories(String... categories) {
        super.setCategories(categories);
        return this;
    }

    @Override
    public LineChart addSeries(String name, Double... values) {
        super.addSeries(name, values);
        return this;
    }

    public static LineChart NEW() {
        return new LineChart();
    }

    private LineChart() {}

}
