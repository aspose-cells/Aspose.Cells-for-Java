package AsposeCellsExamples.Slicers;

import AsposeCellsExamples.Utils;
import com.aspose.cells.*;

public class CreateSlicerToExcelTable {
	
	static String sourceDir = Utils.Get_SourceDirectory();
	static String outputDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {
		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		// ExStart:1
		// Load sample Excel file containing a table.
		Workbook workbook = new Workbook(sourceDir + "sampleCreateSlicerToExcelTable.xlsx");

		// Access first worksheet.
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Access first table inside the worksheet.
		ListObject table = worksheet.getListObjects().get(0);

		// Add slicer
		int idx = worksheet.getSlicers().add(table, 0, "H5");

		// Save the workbook in output XLSX format.
		workbook.save(outputDir + "outputCreateSlicerToExcelTable.xlsx", SaveFormat.XLSX);
		// ExEnd:1

		// Print the message
		System.out.println("CreateSlicerToPivotTable executed successfully.");
	}
}
