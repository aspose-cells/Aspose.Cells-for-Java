package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.PivotTable;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;

public class DeletePivotTableFromWorksheet {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(DeletePivotTableFromWorksheet.class) + "TechnicalArticles/";

		// Create workbook object from source Excel file
		Workbook workbook = new Workbook(dataDir + "sample.xlsx");

		// Access the first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Access the first pivot table object
		PivotTable pivotTable = worksheet.getPivotTables().get(0);

		// Remove pivot table using pivot table object
		worksheet.getPivotTables().remove(pivotTable);

		// Remove pivot table using pivot table position
		worksheet.getPivotTables().removeAt(0);

		// Save the workbook
		workbook.save(dataDir + "DPTableFromWorksheet_out.xlsx");

	}
}
