package AsposeCellsExamples.PivotTables;

import AsposeCellsExamples.Utils;
import com.aspose.cells.*;

public class PivotTableGetExternalConnectionDataSource {
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// Source directory
		String sourceDir = Utils.Get_SourceDirectory();

		// Load sample file
		Workbook workbook = new Workbook(sourceDir + "SamplePivotTableExternalConnection.xlsx");

		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Get the pivot table
		PivotTable pivotTable = worksheet.getPivotTables().get(0);

		// Print External Connection Details
		System.out.println("External Connection Data Source");
		System.out.println("Name: " + pivotTable.getExternalConnectionDataSource().getName());
		System.out.println("Type: " + pivotTable.getExternalConnectionDataSource().getType());
		//ExEnd: 1

		System.out.println("PivotTableGetExternalConnectionDataSource executed successfully.");
	}
}
