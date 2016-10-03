package com.aspose.cells.examples.PivotTables;

import com.aspose.cells.PivotField;
import com.aspose.cells.PivotFieldCollection;
import com.aspose.cells.PivotFieldDataDisplayFormat;
import com.aspose.cells.PivotItemPosition;
import com.aspose.cells.PivotTable;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class SettingDataFieldFormat {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SettingDataFieldFormat.class) + "PivotTables/";
		// Load a template file
		Workbook workbook = new Workbook(dataDir + "PivotTable.xls");

		// Get the first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);
		PivotTable pivotTable = worksheet.getPivotTables().get(0);
		// Accessing the data fields.
		PivotFieldCollection pivotFields = pivotTable.getDataFields();

		// Accessing the first data field in the data fields.
		PivotField pivotField = pivotFields.get(0);

		// Setting data display format
		pivotField.setDataDisplayFormat(PivotFieldDataDisplayFormat.PERCENTAGE_OF);

		// Setting the base field.
		pivotField.setBaseField(1);

		// Setting the base item.
		pivotField.setBaseItem(PivotItemPosition.NEXT);

		// Setting number format
		pivotField.setNumber(10);
	}
}
