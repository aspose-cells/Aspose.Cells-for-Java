package AsposeCellsExamples.Data;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class GetAllHiddenRowsIndicesAfterRefreshingAutoFilter { 
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		//Load the sample Excel file
		Workbook wb = new Workbook(srcDir + "sampleGetAllHiddenRowsIndicesAfterRefreshingAutoFilter.xlsx");
		 
		//Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);
		 
		//Apply autofilter
		ws.getAutoFilter().addFilter(0, "Orange");
		 
		//True means, it will refresh autofilter and return hidden rows.
		//False means, it will not refresh autofilter but return same hidden rows.
		int[] rowIndices = ws.getAutoFilter().refresh(true);
		 
		System.out.println("Printing Rows Indices, Cell Names and Values Hidden By AutoFilter.");
		System.out.println("--------------------------");
		 
		for(int i=0; i<rowIndices.length; i++)
		{
			int r = rowIndices[i];
			Cell cell = ws.getCells().get(r, 0);
			System.out.println(r + "\t" + cell.getName() + "\t" + cell.getStringValue());
		}
			 
		// Print the message
		System.out.println("GetAllHiddenRowsIndicesAfterRefreshingAutoFilter executed successfully.");
	}
}
