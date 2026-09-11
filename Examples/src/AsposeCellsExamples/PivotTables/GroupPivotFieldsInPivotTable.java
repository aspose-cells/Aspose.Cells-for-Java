package AsposeCellsExamples.PivotTables;

import java.util.*;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class GroupPivotFieldsInPivotTable {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		//Load sample workbook
		Workbook wb = new Workbook(srcDir + "sampleGroupPivotFieldsInPivotTable.xlsx");
		 
		//Access the second worksheet
		Worksheet ws = wb.getWorksheets().get(1);
		 
		//Access the pivot table
		PivotTable pt = ws.getPivotTables().get(0);
		 
		//Specify the start and end date time
		DateTime dtStart = new DateTime(2008, 1, 1);//1-Jan-2018
		DateTime dtEnd = new DateTime(2008, 9, 5); //5-Sep-2018
		 
		//Apply the grouping by months and quarters on first pivot field
        pt.getBaseFields().get(0).groupBy(dtStart, dtEnd, new int[] {PivotGroupByType.MONTHS,PivotGroupByType.QUARTERS}, 0, true);
		 
		//Refresh and calculate pivot table
		pt.getPivotCache().refresh();
		 
		//Save the output Excel file
		wb.save(outDir + "outputGroupPivotFieldsInPivotTable.xlsx");

		// Print the message
		System.out.println("GroupPivotFieldsInPivotTable executed successfully.");
	}
}
