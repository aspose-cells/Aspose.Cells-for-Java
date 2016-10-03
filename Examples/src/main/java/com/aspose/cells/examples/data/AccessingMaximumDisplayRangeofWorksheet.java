package com.aspose.cells.examples.data;

import com.aspose.cells.Range;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AccessingMaximumDisplayRangeofWorksheet {
	public static void main(String[] args) throws Exception {
		// Path to source file
		String dataDir = Utils.getSharedDataDir(AccessingMaximumDisplayRangeofWorksheet.class) + "data/";

		// Instantiate a workbook from source file
		Workbook workbook = new Workbook(dataDir + "Book1.xlsx");

		// Access the first workbook
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Access the Maximum Display Range
		Range range = worksheet.getCells().getMaxDisplayRange();

		// Print the Maximum Display Range RefersTo property
		System.out.println("Maximum Display Range: " + range.getRefersTo());
	}
}
