package com.aspose.cells.examples.rows_cloumns;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class DeleteMultipleRows {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(DeleteMultipleRows.class) + "rows_cloumns/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "Book1.xlsx");

		// Accessing the first worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Deleting 10 rows from the worksheet starting from 3rd row
		worksheet.getCells().deleteRows(2, 10, true);

		// Saving the modified Excel file in default (that is Excel 2000) format
		workbook.save(dataDir + "DeleteMultipleRows_out.xls");
	}
}
