package com.aspose.cells.examples.articles;

import com.aspose.cells.AxisType;
import com.aspose.cells.Chart;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class DetermineWhichAxisExistsInChart {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(DetermineWhichAxisExistsInChart.class) + "articles/";
		// Create workbook object
		Workbook workbook = new Workbook(dataDir + "source.xlsx");

		// Access the first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Access the chart
		Chart chart = worksheet.getCharts().get(0);

		// Determine which axis exists in chart
		boolean ret = chart.hasAxis(AxisType.CATEGORY, true);
		System.out.println("Has Primary Category Axis: " + ret);

		ret = chart.hasAxis(AxisType.CATEGORY, false);
		System.out.println("Has Secondary Category Axis: " + ret);

		ret = chart.hasAxis(AxisType.VALUE, true);
		System.out.println("Has Primary Value Axis: " + ret);

		ret = chart.hasAxis(AxisType.VALUE, false);
		System.out.println("Has Seconary Value Axis: " + ret);

	}
}
