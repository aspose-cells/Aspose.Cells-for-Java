package com.aspose.cells.examples.charts;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class UsingSparklines {

	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(UsingSparklines.class) + "charts/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook();
		WorksheetCollection worksheets = workbook.getWorksheets();

		// Obtaining the reference of the first worksheet
		Worksheet worksheet = worksheets.get(0);
		Cells cells = worksheet.getCells();

		System.out.println("Sparkline count: " + worksheet.getSparklineGroupCollection().getCount());

		for (int i = 0; i < worksheet.getSparklineGroupCollection().getCount(); i++) {
			SparklineGroup g = worksheet.getSparklineGroupCollection().get(i);
			System.out.println("sparkline group: type:" + g.getType());

			for (int j = 0; j < g.getSparklineCollection().getCount(); i++) {
				Sparkline gg = g.getSparklineCollection().get(i);
				System.out.println("sparkline: row:" + gg.getRow() + ", col:" + gg.getColumn() + ", dataRange:"
						+ gg.getDataRange());
			}
		}
		// Add Sparklines
		// Define the CellArea D2:D10
		CellArea ca = new CellArea();
		ca.StartColumn = 4;
		ca.EndColumn = 4;
		ca.StartRow = 1;
		ca.EndRow = 7;
		int idx = worksheet.getSparklineGroupCollection().add(SparklineType.COLUMN, "Sheet1!B2:D8", false, ca);

		SparklineGroup group = worksheet.getSparklineGroupCollection().get(idx);
		// Create CellsColor
		CellsColor clr = workbook.createCellsColor();
		clr.setColor(Color.getChocolate());
		group.setSeriesColor(clr);
		workbook.save(dataDir + "UsingSparklines_out.xls");

		// Print message
		System.out.println("Workbook with chart is created successfully.");
	}
}
