package com.aspose.cells.examples.articles;

import com.aspose.cells.Chart;
import com.aspose.cells.ChartType;
import com.aspose.cells.SheetType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class CreatePivotChartbasedonPivotTable {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CreatePivotChartbasedonPivotTable.class) + "articles/";
		// Instantiating an Workbook object
		Workbook workbook = new Workbook(dataDir + "pivotTable_test.xls");
		// Adding a new sheet
		int sheetIndex = workbook.getWorksheets().add(SheetType.CHART);
		Worksheet sheet3 = workbook.getWorksheets().get(sheetIndex);
		// Naming the sheet
		sheet3.setName("PivotChart");
		// Adding a column chart
		int chartIndex = sheet3.getCharts().add(ChartType.COLUMN, 0, 5, 28, 16);
		Chart chart = sheet3.getCharts().get(chartIndex);
		// Setting the pivot chart data source
		chart.setPivotSource("PivotTable!PivotTable1");
		chart.setHidePivotFieldButtons(false);
		// Saving the Excel file
		workbook.save(dataDir + "CPCBasedOnPTable_out.xls");

	}
}
