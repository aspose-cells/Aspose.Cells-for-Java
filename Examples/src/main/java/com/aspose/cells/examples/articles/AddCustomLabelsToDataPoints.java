package com.aspose.cells.examples.articles;

import com.aspose.cells.Chart;
import com.aspose.cells.ChartPoint;
import com.aspose.cells.ChartType;
import com.aspose.cells.FileFormatType;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Series;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AddCustomLabelsToDataPoints {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddCustomLabelsToDataPoints.class) + "articles/";
		Workbook workbook = new Workbook(FileFormatType.XLSX);
		Worksheet sheet = workbook.getWorksheets().get(0);

		// Put data
		sheet.getCells().get(0, 0).putValue(1);
		sheet.getCells().get(0, 1).putValue(2);
		sheet.getCells().get(0, 2).putValue(3);

		sheet.getCells().get(1, 0).putValue(4);
		sheet.getCells().get(1, 1).putValue(5);
		sheet.getCells().get(1, 2).putValue(6);

		sheet.getCells().get(2, 0).putValue(7);
		sheet.getCells().get(2, 1).putValue(8);
		sheet.getCells().get(2, 2).putValue(9);

		// Generate the chart
		int chartIndex = sheet.getCharts().add(ChartType.SCATTER_CONNECTED_BY_LINES_WITH_DATA_MARKER, 5, 1, 24, 10);
		Chart chart = sheet.getCharts().get(chartIndex);

		chart.getTitle().setText("Test");
		chart.getCategoryAxis().getTitle().setText("X-Axis");
		chart.getValueAxis().getTitle().setText("Y-Axis");

		chart.getNSeries().setCategoryData("A1:C1");

		// Insert series
		chart.getNSeries().add("A2:C2", false);

		Series series = chart.getNSeries().get(0);

		int pointCount = series.getPoints().getCount();
		for (int i = 0; i < pointCount; i++) {
			ChartPoint pointIndex = series.getPoints().get(i);

			pointIndex.getDataLabels().setText("Series 1" + "\n" + "Point " + i);
		}

		// Insert series
		chart.getNSeries().add("A3:C3", false);

		series = chart.getNSeries().get(1);

		pointCount = series.getPoints().getCount();
		for (int i = 0; i < pointCount; i++) {
			ChartPoint pointIndex = series.getPoints().get(i);

			pointIndex.getDataLabels().setText("Series 2" + "\n" + "Point " + i);
		}

		workbook.save(dataDir + "ACLToDataPoints_out.xlsx", SaveFormat.XLSX);

	}
}
