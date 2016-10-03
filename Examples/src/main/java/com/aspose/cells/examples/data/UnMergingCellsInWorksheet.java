package com.aspose.cells.examples.data;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class UnMergingCellsInWorksheet {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(UnMergingCellsInWorksheet.class) + "data/";

		// Create a Workbook.
		Workbook wbk = new Workbook(dataDir + "mergingcells.xls");

		// Create a Worksheet and get the first sheet.
		Worksheet worksheet = wbk.getWorksheets().get(0);

		// Create a Cells object to fetch all the cells.
		Cells cells = worksheet.getCells();

		// Unmerge the cells.
		cells.unMerge(5, 2, 2, 3);

		// Save the file.
		wbk.save(dataDir + "UnMergingCellsInWorksheet_out.xls");

		// Print message
		System.out.println("Process completed successfully");

	}
}
