package com.aspose.cells.examples.charts;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class SettingChartsData {

	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SettingChartsData.class) + "charts/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook();
		WorksheetCollection worksheets = workbook.getWorksheets();

		// Obtaining the reference of the first worksheet
		Worksheet worksheet = worksheets.get(0);
		Cells cells = worksheet.getCells();

		// Adding a sample value to "A1" cell
		cells.get("A1").setValue(50);

		// Adding a sample value to "A2" cell
		cells.get("A2").setValue(100);

		// Adding a sample value to "A3" cell
		cells.get("A3").setValue(150);

		// Adding a sample value to "A4" cell
		cells.get("A4").setValue(200);

		// Adding a sample value to "B1" cell
		cells.get("B1").setValue(60);

		// Adding a sample value to "B2" cell
		cells.get("B2").setValue(32);

		// Adding a sample value to "B3" cell
		cells.get("B3").setValue(50);

		// Adding a sample value to "B4" cell
		cells.get("B4").setValue(40);

		// Adding a sample value to "C1" cell as category data
		cells.get("C1").setValue("Q1");

		// Adding a sample value to "C2" cell as category data
		cells.get("C2").setValue("Q2");

		// Adding a sample value to "C3" cell as category data
		cells.get("C3").setValue("Y1");

		// Adding a sample value to "C4" cell as category data
		cells.get("C4").setValue("Y2");

		// Adding a chart to the worksheet
		ChartCollection charts = worksheet.getCharts();

		// Accessing the instance of the newly added chart
		int chartIndex = charts.add(ChartType.COLUMN, 5, 0, 15, 5);
		Chart chart = charts.get(chartIndex);

		// Adding NSeries (chart data source) to the chart ranging from "A1"
		// cell to "B4"
		SeriesCollection nSeries = chart.getNSeries();
		nSeries.add("A1:B4", true);

		// Setting the data source for the category data of NSeries
		nSeries.setCategoryData("C1:C4");

		workbook.save(dataDir + "SettingChartsData_out.xls");

		// Print message
		System.out.println("Workbook with chart is created successfully.");
	}
}
