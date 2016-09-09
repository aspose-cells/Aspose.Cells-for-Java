package com.aspose.cells.examples.charts.ManipulateChart;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class ModifyPieChart {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ModifyPieChart.class) + "charts/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "piechart.xls");

		// Obtaining the reference of the first worksheet
		WorksheetCollection worksheets = workbook.getWorksheets();
		Worksheet sheet = worksheets.get(1);

		// Load the chart from source worksheet
		Chart chart = sheet.getCharts().get(0);
		DataLabels datalabels = chart.getNSeries().get(0).getPoints().get(0).getDataLabels();
		datalabels.setText("aspose");

		// Saving the Excel file
		workbook.save(dataDir + "MPieChart-out.xls");

		// Print message
		System.out.println("Line chart is successfully modified.");
		// ExEnd:1

	}
}
