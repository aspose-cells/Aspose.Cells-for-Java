package com.aspose.cells.examples.worksheets;

import com.aspose.cells.PageSetup;
import com.aspose.cells.PrintCommentsType;
import com.aspose.cells.PrintErrorsType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

public class OtherPrintOptions {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(OtherPrintOptions.class) + "worksheets/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Accessing the first worksheet in the Workbook file
		WorksheetCollection worksheets = workbook.getWorksheets();
		Worksheet sheet = worksheets.get(0);

		// Obtaining the reference of the PageSetup of the worksheet
		PageSetup pageSetup = sheet.getPageSetup();

		// Allowing to print gridlines
		pageSetup.setPrintGridlines(true);

		// Allowing to print row/column headings
		pageSetup.setPrintHeadings(true);

		// Allowing to print worksheet in black & white mode
		pageSetup.setBlackAndWhite(true);

		// Allowing to print comments as displayed on worksheet
		pageSetup.setPrintComments(PrintCommentsType.PRINT_IN_PLACE);

		// Allowing to print worksheet with draft quality
		pageSetup.setPrintDraft(true);

		// Allowing to print cell errors as N/A
		pageSetup.setPrintErrors(PrintErrorsType.PRINT_ERRORS_NA);
		workbook.save(dataDir + "OtherPrintOptions_out.xls");
	}
}
