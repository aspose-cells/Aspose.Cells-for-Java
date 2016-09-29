package com.aspose.cells.examples.tables;

import com.aspose.cells.ListObjectCollection;
import com.aspose.cells.TotalsCalculation;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class CreatingListObject {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CreatingListObject.class) + "tables/";
		// Create a Workbook object.
		// Open a template excel file.
		Workbook workbook = new Workbook(dataDir + "book1.xlsx");

		// Get the List objects collection in the first worksheet.
		ListObjectCollection listObjects = workbook.getWorksheets().get(0).getListObjects();

		// Add a List based on the data source range with headers on.
		listObjects.add(1, 1, 11, 5, true);

		// Show the total row for the List.
		listObjects.get(0).setShowTotals(true);

		// Calculate the total of the last (5th ) list column.
		listObjects.get(0).getListColumns().get(4).setTotalsCalculation(TotalsCalculation.SUM);

		// Save the excel file.
		workbook.save(dataDir + "CreatingListObject_out.xls");
	}
}
