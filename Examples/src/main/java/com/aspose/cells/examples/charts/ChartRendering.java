package com.aspose.cells.examples.charts;

import java.awt.RenderingHints;

import com.aspose.cells.Chart;
import com.aspose.cells.ChartCollection;
import com.aspose.cells.ChartType;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

public class ChartRendering {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(CreateChart.class) + "charts/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Obtaining the reference of the first worksheet
		WorksheetCollection worksheets = workbook.getWorksheets();

		Worksheet sheet = worksheets.get(0);
		ChartCollection charts = sheet.getCharts();

		// Adding a chart to the worksheet
		int chartIndex = charts.add(ChartType.PYRAMID, 5, 0, 15, 5);
		Chart chart = charts.get(chartIndex);

		// Create an instance of ImageOrPrintOptions and set a few properties
		ImageOrPrintOptions options = new ImageOrPrintOptions();
		options.setVerticalResolution(300);
		options.setHorizontalResolution(300);
		options.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
		options.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);

		// Convert chart to image with additional settings
		chart.toImage(dataDir + "chart.png", options);
	}
}
