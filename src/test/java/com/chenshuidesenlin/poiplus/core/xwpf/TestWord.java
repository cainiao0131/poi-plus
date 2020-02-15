package com.chenshuidesenlin.poiplus.core.xwpf;

import com.chenshuidesenlin.poiplus.core.xwpf.chart.HistogramChart;
import com.chenshuidesenlin.poiplus.core.xwpf.chart.LineChart;
import com.chenshuidesenlin.poiplus.core.xwpf.entity.Word;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class TestWord {

    @Test
    public void testAddChart() {
        try (Word word = Word.NEW();
             OutputStream out = new FileOutputStream("chart-output.docx"))
        {
            LineChart lineChart = LineChart.NEW();
            lineChart.setTitle("流量安全设备事件统计");
            lineChart.setCategories("2020-02-10", "2020-02-11", "2020-02-12", "2020-02-13");
            lineChart.addSeries("数量", 15.0, 20.0, 14.5, 11.2);
            lineChart.addSeries("质量", 21.0, 2.2, 23.4, 25.2);

            HistogramChart histogramChart = HistogramChart.NEW();
            histogramChart.setTitle("test");
            histogramChart.setCategories("2020-02-10", "2020-02-11", "2020-02-12", "2020-02-13");
            histogramChart.addSeries("abc", 15.0, 20.0, 14.5, 11.2);
            histogramChart.addSeries("ddd", 21.0, 2.2, 23.4, 25.2);

            word
                .addChart(lineChart)
                .addChart(histogramChart)
                .write(out);
        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
        }
    }

}
