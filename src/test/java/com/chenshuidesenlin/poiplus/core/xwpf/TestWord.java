package com.chenshuidesenlin.poiplus.core.xwpf;

import com.chenshuidesenlin.poiplus.core.xwpf.chart.HistogramChart;
import com.chenshuidesenlin.poiplus.core.xwpf.chart.LineChart;
import com.chenshuidesenlin.poiplus.core.xwpf.chart.PieChart;
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
             OutputStream out = new FileOutputStream("word-output.docx"))
        {
            PieChart pieChart = PieChart.NEW();
            pieChart.setTitle("流量安全设备事件统计");
            pieChart.setCategories("Web攻击", "入侵事件", "资产识别", "APP应用识别", "DDOS攻击");
            pieChart.addSeries("数量", 1.0, 20.0, 14.0, 10.0, 3.0);

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
                .addChart(pieChart)
                .addChart(lineChart)
                .addChart(histogramChart)
                .addParagraph("testtestestttesttesttesttesttesttesttesttesttes" +
                    "testtesttesttesttesttesttesttesttesttestttesttesttesttesttesttest")
                .write(out);
        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
        }
    }

}
