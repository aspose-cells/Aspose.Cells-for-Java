package com.aspose.cells.examples.worksheets;

import com.aspose.cells.Workbook;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

public class CopyWithinWorkbook {
	public static void main(String[] args) throws Exception {

		String dataDir = Utils.getSharedDataDir(AddingPageBreaks.class) + "worksheets/";
		// Create a new Workbook by excel file path
		Workbook wb = new Workbook(dataDir + "book1.xls");

		// Create a Worksheets object with reference to the sheets of the Workbook.
		WorksheetCollection sheets = wb.getWorksheets();

		// Copy data to a new sheet from an existing sheet within the Workbook.
		sheets.addCopy("Sheet1");

		// Save the excel file.
		wb.save(dataDir + "CopyWithinWorkbook_out.xls");
	}
}
