package AsposeCellsExamples.PivotTables;

import AsposeCellsExamples.Utils;
import com.aspose.cells.*;

public class PivotTableDataDisplayFormatRanking {
	public static void main(String[] args) throws Exception {
		// ExStart: 1
		// directories
		String sourceDir = Utils.Get_SourceDirectory();
		String outputDir = Utils.Get_OutputDirectory();

		// Load a template file
		Workbook workbook = new Workbook(sourceDir + "PivotTableSample.xlsx");

		// Get the first worksheet
		Worksheet sheet = workbook.getWorksheets().get(0);
		int pivotIndex = 0;

		// Get the pivot tables in the sheet
		PivotTable pivotTable = sheet.getPivotTables().get(pivotIndex);

		// Accessing the data fields.
		PivotFieldCollection pivotFields = pivotTable.getDataFields();

		// Accessing the first data field in the data fields.
		PivotField pivotField = pivotFields.get(0);

		// Setting data display format
		pivotField.setDataDisplayFormat(PivotFieldDataDisplayFormat.RANK_LARGEST_TO_SMALLEST);

		pivotTable.calculateData();
		// Saving the Excel file
		workbook.save(outputDir + "PivotTableDataDisplayFormatRanking_out.xlsx");
		// ExEnd:1

		System.out.println("PivotTableDataDisplayFormatRanking executed successfully.");
	}
}
