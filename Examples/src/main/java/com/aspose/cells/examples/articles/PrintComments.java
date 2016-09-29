package com.aspose.cells.examples.articles;

import com.aspose.cells.PrintCommentsType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class PrintComments {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(PrintComments.class) + "articles/";
		// Create a workbook from source Excel file
		Workbook workbook = new Workbook(dataDir + "source.xlsx");

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Print no comments
		worksheet.getPageSetup().setPrintComments(PrintCommentsType.PRINT_NO_COMMENTS);

		// Print the comments as displayed on sheet
		worksheet.getPageSetup().setPrintComments(PrintCommentsType.PRINT_IN_PLACE);

		// Print the comments at the end of sheet
		worksheet.getPageSetup().setPrintComments(PrintCommentsType.PRINT_SHEET_END);

		// Save workbook in pdf format
		workbook.save(dataDir + "PrintComments_out.pdf");

	}
}
