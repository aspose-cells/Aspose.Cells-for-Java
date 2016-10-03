package com.aspose.cells.examples.PivotTables;

import com.aspose.cells.PivotField;
import com.aspose.cells.PivotFieldCollection;
import com.aspose.cells.PivotFieldSubtotalType;
import com.aspose.cells.PivotTable;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class SetRowColumnPageFieldsFormat {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SetRowColumnPageFieldsFormat.class) + "PivotTables/";
		// Load a template file
		Workbook workbook = new Workbook(dataDir + "PivotTable.xls");

		// Get the first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);
		PivotTable pivotTable = worksheet.getPivotTables().get(0);
		// Accessing the row fields.
		PivotFieldCollection pivotFields = pivotTable.getRowFields();

		// Accessing the first row field in the row fields.
		PivotField pivotField = pivotFields.get(0);

		// Setting Subtotals.
		pivotField.setSubtotals(PivotFieldSubtotalType.SUM, true);
		pivotField.setSubtotals(PivotFieldSubtotalType.COUNT, true);

		// Setting autosort options. Setting the field auto sort.
		pivotField.setAutoSort(true);

		// Setting the field auto sort ascend.
		pivotField.setAscendSort(true);

		// Setting the field auto sort using the field itself.
		pivotField.setAutoSortField(-1);

		// Setting autoShow options. Setting the field auto show.
		pivotField.setAutoShow(true);

		// Setting the field auto show ascend.
		pivotField.setAscendShow(false);

		// Setting the auto show using field(data field).
		pivotField.setAutoShowField(0);
	}
}
