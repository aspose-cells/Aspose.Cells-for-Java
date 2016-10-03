package com.aspose.cells.examples.rows_cloumns;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class InsertingARow {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(InsertingARow.class) + "rows_cloumns/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Accessing the first worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Inserting a row into the worksheet at 3rd position
		worksheet.getCells().insertRows(2, 1);

		// Saving the modified Excel file in default (that is Excel 2000) format
		workbook.save(dataDir + "InsertingARow_out.xls");
	}
}
