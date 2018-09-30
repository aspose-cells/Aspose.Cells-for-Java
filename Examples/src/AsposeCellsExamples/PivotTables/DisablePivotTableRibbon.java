package AsposeCellsExamples.PivotTables;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.PivotTable;
import com.aspose.cells.Range;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;

public class DisablePivotTableRibbon {
	
	// The path to the documents directory.
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		// ExStart:DisablePivotTableRibbon
		// Open the template file containing the pivot table
		Workbook wb = new Workbook(srcDir + "pivot_table_test.xlsx");

		// Access the pivot table in the first sheet
		PivotTable pt = wb.getWorksheets().get(0).getPivotTables().get(0);

		// Disable ribbon for this pivot table
		pt.setEnableWizard(false);

		// Save output file
		wb.save(outDir + "out_java.xlsx");	
		
		// Print the message
		System.out.println("Disable Pivot Table Ribbon executed successfully.");
		
		// ExEnd:DisablePivotTableRibbon	
	}
}
