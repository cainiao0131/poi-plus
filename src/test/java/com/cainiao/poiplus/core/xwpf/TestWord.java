package com.cainiao.poiplus.core.xwpf;

import com.cainiao.poiplus.core.xwpf.dto.chart.HistogramChart;
import com.cainiao.poiplus.core.xwpf.dto.chart.LineChart;
import com.cainiao.poiplus.core.xwpf.mylib.Word;
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
            word
                .addChart(LineChart.NEW()
                    .setTitle("流量安全设备事件统计")
                    .setCategories("2020-02-10", "2020-02-11", "2020-02-12", "2020-02-13")
                    .addSeries("数量", 15.0, 20.0, 14.5, 11.2)
                    .addSeries("质量", 21.0, 2.2, 23.4, 25.2))
                .addChart(HistogramChart.NEW()
                    .setTitle("test")
                    .setCategories("2020-02-10", "2020-02-11", "2020-02-12", "2020-02-13")
                    .addSeries("abc", 15.0, 20.0, 14.5, 11.2)
                    .addSeries("ddd", 21.0, 2.2, 23.4, 25.2))
                .write(out);
        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
        }
    }

}
