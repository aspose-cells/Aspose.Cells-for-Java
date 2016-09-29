package com.aspose.cells.examples.articles;

import com.aspose.cells.QueryTable;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ReadingAndWritingQueryTable {
	public static void main(String[] args) throws Exception {

		String dataDir = Utils.getSharedDataDir(ReadingAndWritingQueryTable.class) + "articles/";
		// Create workbook from source excel file
		Workbook workbook = new Workbook(dataDir + "Sample.xlsx");

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Access first Query Table
		QueryTable qt = worksheet.getQueryTables().get(0);

		// Print Query Table Data
		System.out.println("Adjust Column Width: " + qt.getAdjustColumnWidth());
		System.out.println("Preserve Formatting: " + qt.getPreserveFormatting());

		// Now set Preserve Formatting to true
		qt.setPreserveFormatting(true);

		// Save the workbook
		workbook.save(dataDir + "RAWQueryTable_out.xlsx");


	}
}
