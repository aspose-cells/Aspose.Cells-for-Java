package com.aspose.cells.examples.charts;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class SettingTitlesAxes {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SettingTitlesAxes.class) + "charts/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");
		WorksheetCollection worksheets = workbook.getWorksheets();

		// Obtaining the reference of the newly added worksheet by passing its
		// sheet index
		Worksheet worksheet = worksheets.get(0);
		Cells cells = worksheet.getCells();

		// Adding a chart to the worksheet
		ChartCollection charts = worksheet.getCharts();

		// Accessing the instance of the newly added chart
		int chartIndex = charts.add(ChartType.COLUMN, 5, 0, 15, 7);
		Chart chart = charts.get(chartIndex);

		// Setting the title of a chart
		Title title = chart.getTitle();
		title.setText("ASPOSE");

		// Setting the font color of the chart title to blue
		Font font = title.getFont();
		font.setColor(Color.getBlue());

		// Setting the title of category axis of the chart
		Axis categoryAxis = chart.getCategoryAxis();
		title = categoryAxis.getTitle();
		title.setText("Category");

		// Setting the title of value axis of the chart
		Axis valueAxis = chart.getValueAxis();
		title = valueAxis.getTitle();
		title.setText("Value");

		// Adding NSeries (chart data source) to the chart ranging from "A1"
		// cell
		SeriesCollection nSeries = chart.getNSeries();
		nSeries.add("A1:B3", true);

		// Setting the foreground color of the plot area
		ChartFrame plotArea = chart.getPlotArea();
		Area area = plotArea.getArea();
		area.setForegroundColor(Color.getBlue());

		// Setting the foreground color of the chart area
		ChartArea chartArea = chart.getChartArea();
		area = chartArea.getArea();
		area.setForegroundColor(Color.getYellow());

		// Setting the foreground color of the 1st NSeries area
		Series aSeries = nSeries.get(0);
		area = aSeries.getArea();
		area.setForegroundColor(Color.getRed());

		// Setting the foreground color of the area of the 1st NSeries point
		ChartPointCollection chartPoints = aSeries.getPoints();
		ChartPoint point = chartPoints.get(0);
		point.getArea().setForegroundColor(Color.getCyan());

		// Save the Excel file
		workbook.save(dataDir + "SettingTitlesAxes_out.xls");

		// Print message
		System.out.println("Chart Title is changed successfully.");

	}
}
