package AsposeCellsExamples.PivotTables;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.PivotTable;
import com.aspose.cells.Range;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;

public class GetPivotTableRefreshDate {
	// The path to the documents directory.
	static String srcDir = Utils.Get_SourceDirectory();

	public static void main(String[] args) throws Exception {

		// ExStart:1
		// Instantiating a Workbook object
		Workbook workbook = new Workbook(srcDir + "sourcePivotTable.xlsx");

        // Access first worksheet
        Worksheet worksheet = workbook.getWorksheets().get(0);

        // Access first pivot table inside the worksheet
        PivotTable pivotTable = worksheet.getPivotTables().get(0);
        ;

        // Access pivot table refresh by who property
        System.out.println("Pivot table refresh by who = " + pivotTable.getRefreshedByWho());

        // Access pivot table refresh date
        System.out.println("Pivot table refresh date = " + pivotTable.getRefreshDate());
        // ExEnd:1
        
     	// Print message
     	System.out.println("GetPivotTableRefreshDate executed successfully");        
	}
}

