package com.aspose.cells.examples.articles;

import com.aspose.cells.Chart;
import com.aspose.cells.ChartCollection;
import com.aspose.cells.DataLabels;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ResizeChartDataLabelShapeToFitText {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ResizeChartDataLabelShapeToFitText.class) + "articles/";
		// Create an instance of Workbook containing the Chart
		Workbook book = new Workbook(dataDir + "report.xlsx");

		// Access the Worksheet that contains the Chart
		Worksheet sheet = book.getWorksheets().get(0);

		// Access ChartCollection from Worksheet
		ChartCollection charts = sheet.getCharts();

		// Loop over each chart in collection
		for (int chartIndex = 0; chartIndex < charts.getCount(); chartIndex++) {
			// Access indexed chart from the collection
			Chart chart = charts.get(chartIndex);

			for (int seriesIndex = 0; seriesIndex < chart.getNSeries().getCount(); seriesIndex++) {
				// Access the DataLabels of indexed NSeries
				DataLabels labels = chart.getNSeries().get(seriesIndex).getDataLabels();

				// Set ResizeShapeToFitText property to true
				labels.setResizeShapeToFitText(true);
			}

			// Calculate Chart
			chart.calculate();
		}

		// Save the result
		book.save(dataDir + "RCDLabelShapeToFitText_out.xlsx");

	}
}
