package com.aspose.cells.examples.charts;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class ModifyLineChart {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ModifyLineChart.class) + "charts/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Obtaining the reference of the first worksheet
		WorksheetCollection worksheets = workbook.getWorksheets();
		Worksheet sheet = worksheets.get(0);

		// Load the chart from source worksheet
		Chart chart = sheet.getCharts().get(0);

		// Adding NSeries (chart data source) to the chart ranging from "A1"
		// cell to "B3"
		SeriesCollection serieses = chart.getNSeries();
		serieses.add("{20,40,90}", true);
		serieses.add("{110,70,220}", true);

		// Saving the Excel file
		workbook.save(dataDir + "ModifyLineChart_out.xls");

		// Print message
		System.out.println("Line chart is successfully modified.");


	}
}
