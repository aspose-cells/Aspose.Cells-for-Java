package com.aspose.cells.examples.tables;

import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class ConvertTableToRange {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ConvertTableToRange.class) + "tables/";
		// Open an existing file that contains a table/list object in it
		Workbook wb = new Workbook(dataDir + "book1.xlsx");

		// Convert the first table/list object (from the first worksheet) to normal range
		wb.getWorksheets().get(0).getListObjects().get(0).convertToRange();

		// Save the file
		wb.save(dataDir + "ConvertTableToRange_out.xlsx");
	}
}
