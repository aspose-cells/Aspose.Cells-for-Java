package com.aspose.cells.examples.data;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class MakeCellActive {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(MakeCellActive.class) + "data/";
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

		// Set B2 cell as an active cell in the worksheet.
		worksheet1.setActiveCell("B2");

		// Set the B column as the first visible column in the worksheet.
		worksheet1.setFirstVisibleColumn(1);

		// Set the 2nd row as the first visible row in the worksheet.
		worksheet1.setFirstVisibleRow(1);

		// Save the Excel file.
		workbook.save(dataDir + "MakeCellActive_out.xls");
	}
}
