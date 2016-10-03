package com.aspose.cells.examples.PivotTables;

import com.aspose.cells.ConsolidationFunction;
import com.aspose.cells.PivotTable;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ConsolidationFunctions {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ConsolidationFunctions.class) + "PivotTables/";
		// Create workbook from source excel file
		Workbook workbook = new Workbook(dataDir + "sample1.xlsx");

		// Access the first worksheet of the workbook
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Access the first pivot table of the worksheet
		PivotTable pivotTable = worksheet.getPivotTables().get(0);

		// Apply Average consolidation function to first data field
		pivotTable.getDataFields().get(0).setFunction(ConsolidationFunction.AVERAGE);

		// Apply DistinctCount consolidation function to second data field
		pivotTable.getDataFields().get(1).setFunction(ConsolidationFunction.DISTINCT_COUNT);

		// Calculate the data to make changes affect
		pivotTable.calculateData();

		// Save the workbook
		workbook.save(dataDir + "ConsolidationFunctions_out.xlsx");
	}
}
