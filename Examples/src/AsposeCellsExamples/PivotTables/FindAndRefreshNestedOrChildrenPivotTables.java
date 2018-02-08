package AsposeCellsExamples.PivotTables;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class FindAndRefreshNestedOrChildrenPivotTables { 
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		//Load sample Excel file
		Workbook wb = new Workbook(srcDir + "sampleFindAndRefreshNestedOrChildrenPivotTables.xlsx");

		//Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		//Access third pivot table
		PivotTable ptParent = ws.getPivotTables().get(2);

		//Access the children of the parent pivot table
		PivotTable[] ptChildren = ptParent.getChildren();

		//Refresh all the children pivot table
		int count = ptChildren.length;
		for (int idx = 0; idx < count; idx++)
		{
			//Access the child pivot table
			PivotTable ptChild = ptChildren[idx];

			//Refresh the child pivot table
			ptChild.refreshData();
			ptChild.calculateData();
		}
		
		// Print the message
		System.out.println("FindAndRefreshNestedOrChildrenPivotTables executed successfully.");
	}
}
