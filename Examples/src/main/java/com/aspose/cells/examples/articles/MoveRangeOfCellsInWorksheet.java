package com.aspose.cells.examples.articles;

import com.aspose.cells.CellArea;
import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class MoveRangeOfCellsInWorksheet {
	public static void main(String[] args) throws Exception {
		// ExStart:MoveRangeOfCellsInWorksheet
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(MoveRangeOfCellsInWorksheet.class);
		// Instantiate the workbook object. Open the Excel file
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		Cells cells = workbook.getWorksheets().get(0).getCells();

		// Create Cell's area
		CellArea ca = CellArea.createCellArea("A1", "B5");

		// Move Range
		cells.moveRange(ca, 0, 2);

		// Save the resultant file
		workbook.save(dataDir + "book2.xls");
		// ExEnd:MoveRangeOfCellsInWorksheet
	}
}
