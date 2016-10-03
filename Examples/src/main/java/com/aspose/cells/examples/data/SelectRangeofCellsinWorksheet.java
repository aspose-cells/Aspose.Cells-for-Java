package com.aspose.cells.examples.data;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class SelectRangeofCellsinWorksheet {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SelectRangeofCellsinWorksheet.class) + "data/";
		// Instantiate a new Workbook.
		Workbook workbook = new Workbook();

		// Get the first worksheet in the workbook.
		Worksheet worksheet1 = workbook.getWorksheets().get(0);

		// Get the cells in the worksheet.
		Cells cells = worksheet1.getCells();

		// Input data into B2 cell.
		cells.get(1, 1).setValue("Hello World!");

		// Set the first sheet as an active sheet.
		workbook.getWorksheets().setActiveSheetIndex(0);

		// Select range of cells(A1:E10) in the worksheet.
		worksheet1.selectRange(0, 0, 10, 5, true);

		// Save the Excel file.
		workbook.save(dataDir + "SROfCInWorksheet_out.xlsx");
	}
}
