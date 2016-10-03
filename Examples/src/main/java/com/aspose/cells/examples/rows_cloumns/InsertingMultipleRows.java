package com.aspose.cells.examples.rows_cloumns;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class InsertingMultipleRows {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(InsertingMultipleRows.class) + "rows_cloumns/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Accessing the first worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Inserting 10 rows into the worksheet starting from 3rd row
		worksheet.getCells().insertRows(2, 10);

		// Saving the modified Excel file in default (that is Excel 2000) format
		workbook.save(dataDir + "InsertingMultipleRows_out.xls");
	}
}
