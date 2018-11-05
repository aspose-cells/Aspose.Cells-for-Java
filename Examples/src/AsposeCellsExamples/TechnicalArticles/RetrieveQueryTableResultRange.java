package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.Cell;
import com.aspose.cells.FontSetting;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import AsposeCellsExamples.Utils;

public class RetrieveQueryTableResultRange {
	// The path to the documents directory.
	static String srcDir = Utils.Get_SourceDirectory();

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// Create workbook from source excel file
		Workbook wb = new Workbook(srcDir + "Query TXT.xlsx");

		// Display the address(range) of result range of query table
		System.out.println(wb.getWorksheets().get(0).getQueryTables().get(0).getResultRange().getAddress());
		// ExEnd:1
		
		// Print message
		System.out.println("Retrieve Query Table Result Range completed successfully");
	}
}
