package com.aspose.cells.examples.articles;

import com.aspose.cells.Chart;
import com.aspose.cells.ChartPoint;
import com.aspose.cells.Series;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class FindDataPoints {
	public static void main(String[] args) throws Exception {
		// ExStart:FindDataPoints
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(FindDataPoints.class) + "articles/";

		// Load source excel file containing Bar of Pie chart
		Workbook wb = new Workbook(dataDir + "PieBars.xlsx");

		// Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		// Access first chart which is Bar of Pie chart and calculate it
		Chart ch = ws.getCharts().get(0);
		ch.calculate();

		// Access the chart series
		Series srs = ch.getNSeries().get(0);

		// Print the data points of the chart series and check
		// its IsInSecondaryPlot property to determine if data point is inside
		// the bar or pie
		for (int i = 0; i < srs.getPoints().getCount(); i++) {
			// Access chart point
			ChartPoint cp = srs.getPoints().get(i);

			// Skip null values
			if (cp.getYValue() == null)
				continue;

			// Print the chart point value and see if it is inside bar or pie
			// If the IsInSecondaryPlot is true, then the data point is inside
			// bar
			// otherwise it is inside the pie
			System.out.println("Value: " + cp.getYValue());
			System.out.println("IsInSecondaryPlot: " + cp.isInSecondaryPlot());
			System.out.println();
			// ExEnd:FindDataPoints
		}
	}
}
