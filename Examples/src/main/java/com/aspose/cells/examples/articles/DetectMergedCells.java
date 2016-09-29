package com.aspose.cells.examples.articles;

import java.util.ArrayList;

import com.aspose.cells.CellArea;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class DetectMergedCells {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(DetectMergedCells.class) + "articles/";
		// Instantiate a new Workbook
		Workbook wkBook = new Workbook(dataDir + "MergeTrial.xls");
		// Get a worksheet in the workbook
		Worksheet wkSheet = wkBook.getWorksheets().get("Merge Trial");
		// Clear its contents
		wkSheet.getCells().clearContents(0, 0, wkSheet.getCells().getMaxDataRow(),
				wkSheet.getCells().getMaxDataColumn());

		// Create an arraylist object, Get the merged cells list to put it into the arraylist object
		ArrayList<CellArea> al = wkSheet.getCells().getMergedCells();
		// Define cellarea
		CellArea ca;
		// Define some variables
		int frow, fcol, erow, ecol;
		// Loop through the arraylist and get each cellarea to unmerge it
		for (int i = al.size() - 1; i > -1; i--) {
			ca = new CellArea();
			ca = (CellArea) al.get(i);
			frow = ca.StartRow;
			fcol = ca.StartColumn;
			erow = ca.EndRow;
			ecol = ca.EndColumn;
			wkSheet.getCells().unMerge(frow, fcol, erow, ecol);
		}
		// Save the excel file
		wkBook.save(dataDir + "DetectMergedCells_out.xls");

	}
}
