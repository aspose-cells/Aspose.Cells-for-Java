package com.aspose.cells.examples.articles;

import java.util.Iterator;

import com.aspose.cells.Range;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class CheckForShapes {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CheckForShapes.class) + "articles/";

		// Create an instance of Workbook and load an existing spreadsheet
		Workbook book = new Workbook(dataDir + "sample.xlsx");
		// Loop over all worksheets in the workbook
		for (int i = 0; i < book.getWorksheets().getCount(); i++) {
			Worksheet sheet = book.getWorksheets().get(i);
			// Check if worksheet has populated cells
			if (sheet.getCells().getMaxDataRow() != -1) {
				System.out.println(sheet.getName() + " is not empty because one or more cells are populated");
			}
			// Check if worksheet has shapes
			else if (sheet.getShapes().getCount() > 0) {
				System.out.println(sheet.getName() + " is not empty because there are one or more shapes");
			}
			// Check if worksheet has empty initialized cells
			else {
				Range range = sheet.getCells().getMaxDisplayRange();
				Iterator rangeIterator = range.iterator();
				if (rangeIterator.hasNext()) {
					System.out.println(sheet.getName() + " is not empty because one or more cells are initialized");
				}
			}
		}

	}
}
