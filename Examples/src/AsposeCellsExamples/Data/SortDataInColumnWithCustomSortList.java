package AsposeCellsExamples.Data;

import java.util.*;
import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class SortDataInColumnWithCustomSortList { 
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		//Load the source Excel file
		Workbook wb = new Workbook(srcDir + "sampleSortData_CustomSortList.xlsx");
		 
		//Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);
		 
		//Specify cell area - sort from A1 to A40
		CellArea ca = CellArea.createCellArea("A1", "A40");
		 
		//Create Custom Sort list
		String[] customSortList = new String[] { "USA,US", "Brazil,BR", "China,CN", "Russia,RU", "Canada,CA" };
		 
		//Add Key for Column A, Sort it in Ascending Order with Custom Sort List
		wb.getDataSorter().addKey(0, SortOrder.ASCENDING, customSortList);
		wb.getDataSorter().sort(ws.getCells(), ca);
		 
		//Save the output Excel file
		wb.save(outDir + "outputSortData_CustomSortList.xlsx");

		// Print the message
		System.out.println("SortDataInColumnWithCustomSortList executed successfully.");
	}
}
