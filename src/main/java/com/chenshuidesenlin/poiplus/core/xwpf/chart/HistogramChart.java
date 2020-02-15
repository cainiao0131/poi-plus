package com.chenshuidesenlin.poiplus.core.xwpf.chart;

/**
 * 柱状图，x轴表示分类，y轴表示值（x、y轴反过来的横向柱状图的类为：BarChart）
 */
public class HistogramChart extends AbstractChart {

    public static HistogramChart NEW() {
        return new HistogramChart();
    }

    private HistogramChart() {}

}
