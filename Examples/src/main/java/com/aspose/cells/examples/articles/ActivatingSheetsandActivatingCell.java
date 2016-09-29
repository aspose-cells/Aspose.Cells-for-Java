package com.aspose.cells.examples.articles;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ActivatingSheetsandActivatingCell {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ActivatingSheetsandActivatingCell.class) + "articles/";
		// Instantiate a new Workbook
		Workbook workbook = new Workbook();
		// Get the first worksheet in the workbook
		Worksheet worksheet = workbook.getWorksheets().get(0);
		// Get the cells in the worksheet
		Cells cells = worksheet.getCells();
		// Input data into B2 cell
		cells.get(1, 1).putValue("Hello World!");
		// Set the first sheet as an active sheet
		workbook.getWorksheets().setActiveSheetIndex(0);
		// Set B2 cell as an active cell in the worksheet
		worksheet.setActiveCell("B2");
		// Set the B column as the first visible column in the worksheet
		worksheet.setFirstVisibleColumn(1);
		// Set the 2nd row as the first visible row in the worksheet
		worksheet.setFirstVisibleRow(1);
		// Save the excel file
		workbook.save(dataDir + "ASAActivatingCell_out.xls");

	}
}
