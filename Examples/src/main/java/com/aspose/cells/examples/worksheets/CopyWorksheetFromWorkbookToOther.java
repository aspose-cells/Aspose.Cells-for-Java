package com.aspose.cells.examples.worksheets;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.PageSetup;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class CopyWorksheetFromWorkbookToOther {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(AddingPageBreaks.class) + "worksheets/";
		// Create a new Workbook.
		Workbook excelWorkbook0 = new Workbook();

		// Get the first worksheet in the book.
		Worksheet ws0 = excelWorkbook0.getWorksheets().get(0);

		// Put some data into header rows (A1:A4)
		for (int i = 0; i < 5; i++) {
			ws0.getCells().get(i, 0).setValue("Header Row " + i);
		}

		// Put some detail data (A5:A999)
		for (int i = 5; i < 1000; i++) {
			ws0.getCells().get(i, 0).setValue("Detail Row " + i);
		}

		// Define a pagesetup object based on the first worksheet.
		PageSetup pagesetup = ws0.getPageSetup();

		// The first five rows are repeated in each page... It can be seen in print preview.
		pagesetup.setPrintTitleRows("$1:$5");

		// Create another Workbook.
		Workbook excelWorkbook1 = new Workbook();

		// Get the first worksheet in the book.
		Worksheet ws1 = excelWorkbook1.getWorksheets().get(0);

		// Name the worksheet.
		ws1.setName("Sheet1");

		// Copy data from the first worksheet of the first workbook into the first worksheet of the second workbook.
		ws1.copy(ws0);

		// Save the excel file.
		excelWorkbook1.save(dataDir + "CopyWorksheetFromWorkbookToOther_out.xls", FileFormatType.EXCEL_97_TO_2003);
	}
}
