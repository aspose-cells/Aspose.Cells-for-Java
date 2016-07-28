package com.aspose.cells.examples.charts.fundamental;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class ChangeChartPositionAndSize {

	public static void main(String[] args) throws Exception {
		// ExStart:ChangeChartPositionAndSize
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(ChangeChartPositionAndSize.class);

		String filePath = dataDir + "book1.xls";

		Workbook workbook = new Workbook(filePath);

		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Load the chart from source worksheet
		Chart chart = worksheet.getCharts().get(0);

		// Resize the chart
		chart.getChartObject().setWidth(400);
		chart.getChartObject().setHeight(300);

		// Reposition the chart
		chart.getChartObject().setX(250);
		chart.getChartObject().setY(150);

		// Output the file
		workbook.save(dataDir + "book1.out.xls");

		// Print message
		System.out.println("Position and Size of Chart is changed successfully.");
		// ExEnd:ChangeChartPositionAndSize
	}
}
