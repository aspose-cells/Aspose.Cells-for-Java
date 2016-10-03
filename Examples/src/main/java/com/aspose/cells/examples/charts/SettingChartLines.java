package com.aspose.cells.examples.charts;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class SettingChartLines {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SettingChartLines.class) + "charts/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");
		WorksheetCollection worksheets = workbook.getWorksheets();

		// Obtaining the reference of the newly added worksheet by passing its
		// sheet index
		Worksheet worksheet = worksheets.get(0);
		Cells cells = worksheet.getCells();

		// Adding a chart to the worksheet
		ChartCollection charts = worksheet.getCharts();

		Chart chart = charts.get(0);

		// Adding NSeries (chart data source) to the chart ranging from "A1"
		// cell
		SeriesCollection nSeries = chart.getNSeries();
		nSeries.add("A1:B3", true);

		Series aSeries = nSeries.get(0);
		Line line = aSeries.getSeriesLines();
		line.setStyle(LineType.DOT);

		// Applying a triangular marker style on the data markers of an NSeries
		aSeries.getMarker().setMarkerStyle(ChartMarkerType.TRIANGLE);
		// Setting the weight of all lines in an NSeries to medium
		aSeries = nSeries.get(1);
		line = aSeries.getSeriesLines();
		line.setWeight(WeightType.MEDIUM_LINE);

		// Save the Excel file
		workbook.save(dataDir + "SettingChartLines_out.xls");

		// Print message
		System.out.println("ChartArea is settled successfully.");

	}
}
