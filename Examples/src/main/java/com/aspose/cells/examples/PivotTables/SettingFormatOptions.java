package com.aspose.cells.examples.PivotTables;

import com.aspose.cells.PivotFieldType;
import com.aspose.cells.PivotTable;
import com.aspose.cells.PrintOrderType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class SettingFormatOptions {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SettingFormatOptions.class) + "PivotTables/";
		// Load a template file
		Workbook workbook = new Workbook(dataDir + "PivotTable.xls");

		// Get the first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);
		PivotTable pivotTable = worksheet.getPivotTables().get(0);
		// Dragging the third field to the data area.
		pivotTable.addFieldToArea(PivotFieldType.DATA, 2);

		// Show grand totals for rows.
		pivotTable.setRowGrand(true);

		// Show grand totals for columns.
		pivotTable.setColumnGrand(true);

		// Display a custom string in cells that contain null values.
		pivotTable.setDisplayNullString(true);
		pivotTable.setNullString("null");

		// Setting the layout
		pivotTable.setPageFieldOrder(PrintOrderType.DOWN_THEN_OVER);
	}
}
