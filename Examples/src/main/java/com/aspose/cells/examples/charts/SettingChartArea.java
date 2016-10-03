package com.aspose.cells.examples.charts;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class SettingChartArea {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SettingChartArea.class) + "charts/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook();
		WorksheetCollection worksheets = workbook.getWorksheets();

		// Obtaining the reference of the newly added worksheet by passing its
		// sheet index
		Worksheet worksheet = worksheets.get(0);
		Cells cells = worksheet.getCells();
		// Adding a sample value to "A1" cell
		cells.get("A1").setValue(50);

		// Adding a sample value to "A2" cell
		cells.get("A2").setValue(100);

		// Adding a sample value to "A3" cell
		cells.get("A3").setValue(150);

		// Adding a sample value to "B1" cell
		cells.get("B1").setValue(60);

		// Adding a sample value to "B2" cell
		cells.get("B2").setValue(32);

		// Adding a sample value to "B3" cell
		cells.get("B3").setValue(50);

		// Adding a chart to the worksheet
		ChartCollection charts = worksheet.getCharts();

		// Accessing the instance of the newly added chart
		int chartIndex = charts.add(ChartType.COLUMN, 5, 0, 15, 7);
		Chart chart = charts.get(chartIndex);

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
		workbook.save(dataDir + "SettingChartArea_out.xls");

		// Print message
		System.out.println("ChartArea is settled successfully.");

	}
}
