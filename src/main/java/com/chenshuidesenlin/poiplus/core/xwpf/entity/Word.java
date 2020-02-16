package com.chenshuidesenlin.poiplus.core.xwpf.entity;

import com.chenshuidesenlin.poiplus.core.xwpf.chart.Chart;
import com.chenshuidesenlin.poiplus.core.xwpf.chart.HistogramChart;
import com.chenshuidesenlin.poiplus.core.xwpf.chart.LineChart;
import com.chenshuidesenlin.poiplus.core.xwpf.chart.PieChart;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.AxisCrossBetween;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.AxisTickMark;
import org.apache.poi.xddf.usermodel.chart.BarDirection;
import org.apache.poi.xddf.usermodel.chart.BarGrouping;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.Grouping;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.XDDFBarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChart;
import org.apache.poi.xddf.usermodel.chart.XDDFChartAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFLineChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.Closeable;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

public class Word implements Closeable {

    private XWPFDocument xwpfDocument;

    public static Word NEW() {
        Word word = new Word();
        word.xwpfDocument = new XWPFDocument();
        return word;
    }

    public void write(OutputStream out) throws IOException {
        xwpfDocument.write(out);
    }

    public Word addChart(Chart chart) throws IOException, InvalidFormatException {
        // 创建图表
        XWPFChart xwpfChart = xwpfDocument
            .createChart(XDDFChart.DEFAULT_WIDTH * 12, XDDFChart.DEFAULT_HEIGHT * 8);
        // 设置分类坐标轴在底部，即x轴表示分类
        XDDFChartAxis bottomAxis = xwpfChart.createCategoryAxis(AxisPosition.BOTTOM);
        // 设置值坐标轴在左边，即y轴表示值
        XDDFValueAxis leftAxis = xwpfChart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
        leftAxis.setMajorTickMark(AxisTickMark.OUT);
        leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);
        if (chart == null) {
            return this;
        }
        List<Series> seriesList = chart.getSeries();
        if (seriesList == null || seriesList.isEmpty()) {
            return this;
        }

        // 设置数据
        XDDFChartData chartData = null;
        if (chart instanceof LineChart) {
            chartData = xwpfChart.createData(ChartTypes.LINE, bottomAxis, leftAxis);
            XDDFLineChartData tempData = (XDDFLineChartData) chartData;
            tempData.setGrouping(Grouping.STANDARD);
        } else if (chart instanceof HistogramChart) {
            chartData = xwpfChart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
            XDDFBarChartData tempData = (XDDFBarChartData) chartData;
            tempData.setBarGrouping(BarGrouping.CLUSTERED);
            tempData.setBarDirection(BarDirection.COL);
        } else if (chart instanceof PieChart) {
            chartData = xwpfChart.createData(ChartTypes.PIE, bottomAxis, leftAxis);
        }
        if (chartData == null) {
            return this;
        }
        /*
          有多个series时，varyColors会为不同的series设置不同的颜色
          只有一个series时，varyColors会为不同的category设置不同颜色
          PieChart本来就只需要一个series，且需要为不同category设置不同颜色
          而非PieChart需要为不同的series设置不同的颜色
          因此，PieChart始终把varyColors设置为true
          而非PieChart，在只有一个series时，需要把varyColors设置为false
         */
        int seriesLength = seriesList.size();
        chartData.setVaryColors(chart instanceof PieChart || seriesLength != 1);
        List<String> categoryList = chart.getCategories();
        final int categoryLength = categoryList == null ? 0 : categoryList.size();
        String[] categories = new String[categoryLength];
        if (categoryLength > 0) {
            for (int i = 0; i < categoryLength; i++) {
                categories[i] = categoryList.get(i);
            }
        }

        final XDDFDataSource<?> categoryDataSource = XDDFDataSourcesFactory.fromArray(
            categories, xwpfChart.formatRange(new CellRangeAddress(1, categoryLength, 0, 0)), 0);
        for (int i = 0; i < seriesLength; i++) {
            Series series = seriesList.get(i);
            List<Double> valueList = series.getValues();
            final int valueLength = valueList == null ? 0 : valueList.size();
            Double[] values = new Double[valueLength];
            if (valueLength > 0) {
                for (int j = 0; j < valueLength; j++) {
                    values[j] = valueList.get(j);
                }
            }
            int col = i + 1;
            final XDDFNumericalDataSource<? extends Number> valuesData = XDDFDataSourcesFactory
                .fromArray(values,
                xwpfChart.formatRange(new CellRangeAddress(1, categoryLength, col, col)), col);
            valuesData.setFormatCode("General");
            XDDFChartData.Series
                barChartDataSeries = chartData.addSeries(categoryDataSource, valuesData);
            String seriesName = series.getName();
            barChartDataSeries.setTitle(seriesName, xwpfChart.setSheetTitle(seriesName, col));
        }
        xwpfChart.plot(chartData);

        // 设置图例
        XDDFChartLegend legend = xwpfChart.getOrAddLegend();
        legend.setPosition(LegendPosition.TOP);
        legend.setOverlay(false);

        xwpfChart.setTitleText(chart.getTitle());
        xwpfChart.setTitleOverlay(false);
        xwpfChart.setAutoTitleDeleted(false);
        return this;
    }

    public Word addParagraph(String text) {
        XWPFParagraph paragraph = xwpfDocument.createParagraph();
        paragraph.setIndentationFirstLine(480);
        XWPFRun xwpfRun = paragraph.createRun();
        xwpfRun.setText(text == null ? "" : text);
        return this;
    }

    private Word() {}

    public XWPFDocument getXwpfDocument() {
        return xwpfDocument;
    }

    public void setXwpfDocument(XWPFDocument xwpfDocument) {
        this.xwpfDocument = xwpfDocument;
    }

    @Override
    public void close() throws IOException {
        if (xwpfDocument != null) {
            xwpfDocument.close();
        }
    }
}
