package com.aspose.cells.examples.rows_cloumns;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class CopyingColumns {
	public static void main(String[] args) throws Exception {
		
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CopyingColumns.class) + "rows_cloumns/";

		// Create a new Workbook.
		Workbook excelWorkbook = new Workbook(dataDir + "book1.xls");

		// Get the first worksheet in the workbook.
		Worksheet wsTemplate = excelWorkbook.getWorksheets().get(0);

		// Copy the first column from the first worksheet of the first workbook into the first worksheet of the second workbook.
		wsTemplate.getCells().copyColumn(wsTemplate.getCells(), 1, 4);

		// Save the excel file.
		excelWorkbook.save(dataDir + "CopyingColumns_out.xls");

		// Print message
		System.out.println("Row and Column copied successfully.");
		
	}
}
