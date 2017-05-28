package com.aspose.cells.examples.data;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class SpecifyingSortWarningWhileSortingData {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SpecifyingSortWarningWhileSortingData.class) + "data/";

		// Create workbook.
		Workbook workbook = new Workbook(dataDir + "sampleSortAsNumber.xlsx");

		// Access first worksheet.
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Create your cell area.
		CellArea ca = CellArea.createCellArea("A1", "A20");

		// Create your sorter.
		DataSorter sorter = workbook.getDataSorter();

		// Find the index, since we want to sort by column A, so we should know
		// the index for sorter.
		int idx = CellsHelper.columnNameToIndex("A");

		// Add key in sorter, it will sort in Ascending order.
		sorter.addKey(idx, SortOrder.ASCENDING);
		sorter.setSortAsNumber(true);

		// Perform sort.
		sorter.sort(worksheet.getCells(), ca);

		// Save the output workbook.
		workbook.save(dataDir + "outputSortAsNumber.xlsx");

		// Print message
		System.out.println("SpecifyingSortWarningWhileSortingData Done Successfully");

	}
}
