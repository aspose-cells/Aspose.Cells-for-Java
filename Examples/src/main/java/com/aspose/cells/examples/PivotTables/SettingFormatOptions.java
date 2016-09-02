package com.aspose.cells.examples.PivotTables;

import com.aspose.cells.PivotFieldType;
import com.aspose.cells.PivotTable;
import com.aspose.cells.PrintOrderType;
import com.aspose.cells.examples.Utils;

public class SettingFormatOptions {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(SettingFormatOptions.class);

		PivotTable pivotTable = new PivotTable();
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
