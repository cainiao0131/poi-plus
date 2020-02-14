package com.cainiao.poiplus.core.xwpf.dto.chart;

/**
 * 柱状图，x轴表示分类，y轴表示值（x、y轴反过来的横向柱状图的类为：BarChart）
 */
public class HistogramChart extends AbstractChart {

    @Override
    public HistogramChart setTitle(String title) {
        super.setTitle(title);
        return this;
    }

    @Override
    public HistogramChart setCategories(String... categories) {
        super.setCategories(categories);
        return this;
    }

    @Override
    public HistogramChart addSeries(String name, Double... values) {
        super.addSeries(name, values);
        return this;
    }

    public static HistogramChart NEW() {
        return new HistogramChart();
    }

    private HistogramChart() {}

}
