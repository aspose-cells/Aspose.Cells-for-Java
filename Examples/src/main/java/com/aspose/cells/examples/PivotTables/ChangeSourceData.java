package com.aspose.cells.examples.PivotTables;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Range;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;
import com.aspose.cells.examples.introduction.CreatingWorkbook;

public class ChangeSourceData {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ChangeSourceData.class) + "PivotTables/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "PivotTable.xls");

		// Accessing the first worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Populating new data to the worksheet cells
		Cells cells = worksheet.getCells();
		Cell cell = cells.get("A9");
		cell.setValue("Golf");
		cell = cells.get("B9");
		cell.setValue("Qtr4");
		cell = cells.get("C9");
		cell.setValue(7000);

		// Changing named range "DataSource"
		Range range = cells.createRange(0, 0, 8, 2);
		range.setName("DataSource");

		// Saving the modified Excel file in default format
		workbook.save(dataDir + "ChangeSourceData_out.xls");
	}
}
