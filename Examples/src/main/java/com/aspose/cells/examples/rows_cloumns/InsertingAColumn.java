package com.aspose.cells.examples.rows_cloumns;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class InsertingAColumn {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(InsertingAColumn.class) + "rows_cloumns/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Accessing the first worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Inserting a column into the worksheet at 2nd position
		worksheet.getCells().insertColumns(1, 1);

		// Saving the modified Excel file in default (that is Excel 2000) format
		workbook.save(dataDir + "InsertingAColumn_out.xls");
	}
}
