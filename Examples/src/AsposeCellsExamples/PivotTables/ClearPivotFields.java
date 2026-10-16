package AsposeCellsExamples.PivotTables;

import com.aspose.cells.PivotFieldType;
import com.aspose.cells.PivotTable;
import com.aspose.cells.PivotTableCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;

public class ClearPivotFields {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ClearPivotFields.class) + "PivotTables/";
		// Load a template file
		Workbook workbook = new Workbook(dataDir + "PivotTable.xls");

		// Get the first worksheet
		Worksheet sheet = workbook.getWorksheets().get(0);

		// Get the pivot tables in the sheet
		PivotTableCollection pivotTables = sheet.getPivotTables();

		// Get the first PivotTable
		PivotTable pivotTable = pivotTables.get(0);

		// Clear all the data fields
		pivotTable.getDataFields().clear();

		// Add new data field
		pivotTable.addFieldToArea(PivotFieldType.DATA, "Betrag Netto FW");


		// Refresh and calculate the pivot table data
		pivotTable.refresh(null);

		// Save the Excel file
		workbook.save(dataDir + "ClearPivotFields_out.xlsx");
	}
}
