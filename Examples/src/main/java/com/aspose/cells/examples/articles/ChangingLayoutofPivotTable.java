package com.aspose.cells.examples.articles;

import com.aspose.cells.PivotTable;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ChangingLayoutofPivotTable {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ChangingLayoutofPivotTable.class) + "articles/";
		// Create workbook object from source excel file
		Workbook workbook = new Workbook(dataDir + "source.xlsx");

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Access first pivot table
		PivotTable pivotTable = worksheet.getPivotTables().get(0);

		// 1 - Show the pivot table in compact form
		pivotTable.showInCompactForm();

		// Refresh the pivot table
		pivotTable.refreshData();
		pivotTable.calculateData();

		// Save the output
		workbook.save("CompactForm.xlsx");

		// 2 - Show the pivot table in outline form
		pivotTable.showInOutlineForm();

		// Refresh the pivot table
		pivotTable.refreshData();
		pivotTable.calculateData();

		// Save the output
		workbook.save("OutlineForm.xlsx");

		// 3 - Show the pivot table in tabular form
		pivotTable.showInTabularForm();

		// Refresh the pivot table
		pivotTable.refreshData();
		pivotTable.calculateData();

		// Save the output
		workbook.save(dataDir + "CLOfPivotTable_out.xlsx");

	}
}
