package com.aspose.cells.examples.data.processing.filteringandvalidation;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class AutofilterData {

	public static void main(String[] args) throws Exception {
		// ExStart:AutofilterData
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(AutofilterData.class);

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Accessing the first worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Creating AutoFilter by giving the cells range
		AutoFilter autoFilter = worksheet.getAutoFilter();
		CellArea area = new CellArea();
		autoFilter.setRange("A1:B1");

		// Saving the modified Excel file
		workbook.save(dataDir + "output.xls");

		// Print message
		System.out.println("Process completed successfully");
		// ExEnd:AutofilterData
	}
}
