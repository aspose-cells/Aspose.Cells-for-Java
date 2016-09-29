package com.aspose.cells.examples.articles;

import com.aspose.cells.PageSetup;
import com.aspose.cells.PrintCommentsType;
import com.aspose.cells.PrintErrorsType;
import com.aspose.cells.PrintOrderType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class SettingPrintoptions {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SettingPrintoptions.class) + "articles/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "PageSetup.xls");

		// Accessing the first worksheet in the Excel file
		Worksheet sheet = workbook.getWorksheets().get(0);

		PageSetup pageSetup = sheet.getPageSetup();

		// Specifying the cells range (from A1 cell to E30 cell) of the print area
		pageSetup.setPrintArea("A1:E30");

		// Defining column numbers A & E as title columns
		pageSetup.setPrintTitleColumns("$A:$E");

		// Defining row numbers 1 & 2 as title rows
		pageSetup.setPrintTitleRows("$1:$2");

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

		// Setting the printing order of the pages to over then down
		pageSetup.setOrder(PrintOrderType.OVER_THEN_DOWN);

		// Save the workbook
		workbook.save(dataDir + "SettingPrintoptions_out.xls");

	}
}
