package com.aspose.cells.examples.rows_cloumns;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class SettingWidthOfColumn {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SettingWidthOfColumn.class) + "rows_cloumns/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Accessing the first worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);
		Cells cells = worksheet.getCells();

		// Setting the width of the second column to 17.5
		cells.setColumnWidth(1, 17.5);

		// Saving the modified Excel file in default (that is Excel 2003) format
		workbook.save(dataDir + "SettingWidthOfColumn_out.xls");
	}
}
