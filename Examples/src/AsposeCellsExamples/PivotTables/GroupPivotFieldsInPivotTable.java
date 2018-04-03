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
		 
		//Specify the group type list, we want to group by months and quarters
		ArrayList groupTypeList = new ArrayList();
		groupTypeList.add(PivotGroupByType.MONTHS);
		groupTypeList.add(PivotGroupByType.QUARTERS);
		 
		//Apply the grouping on first pivot field
		pt.setManualGroupField(0, dtStart, dtEnd, groupTypeList, 1);
		 
		//Refresh and calculate pivot table
		pt.setRefreshDataFlag(true);
		pt.refreshData();
		pt.calculateData();
		pt.setRefreshDataFlag(false);
		 
		//Save the output Excel file
		wb.save(outDir + "outputGroupPivotFieldsInPivotTable.xlsx");

		// Print the message
		System.out.println("GroupPivotFieldsInPivotTable executed successfully.");
	}
}
