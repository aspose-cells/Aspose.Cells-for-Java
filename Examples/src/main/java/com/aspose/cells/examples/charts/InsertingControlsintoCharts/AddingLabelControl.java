package com.aspose.cells.examples.charts.InsertingControlsintoCharts;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;
import com.aspose.cells.examples.charts.fundamental.CreateChart;

public class AddingLabelControl {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddingLabelControl.class) + "charts/";

		String filePath = dataDir + "chart.xls";

		Workbook workbook = new Workbook(filePath);

		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Load the chart from source worksheet
		Chart chart = worksheet.getCharts().get(0);

		Label label = chart.getShapes().addLabelInChart(100, 100, 350, 900);
		label.setText("Write Label here");
		label.setPlacement(PlacementType.FREE_FLOATING);
		label.getFillFormat().setForeColor(Color.getChocolate());

		// Output the file
		workbook.save(dataDir + "ALControl-out.xls");

		// Print message
		System.out.println("Label added to chart successfully.");

	}
}
