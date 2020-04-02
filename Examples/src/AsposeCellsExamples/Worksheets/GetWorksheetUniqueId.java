package AsposeCellsExamples.Worksheets;

import AsposeCellsExamples.Utils;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class GetWorksheetUniqueId {
	public static void main(String[] args) throws Exception {
		// ExStart:1
		//Source directory
		String sourceDir = Utils.Get_SourceDirectory();

		// Load source Excel file
		Workbook workbook = new Workbook(sourceDir + "Book1.xlsx");

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Print Unique Id
		System.out.println("Unique Id: " + worksheet.getUniqueId());
		// ExEnd:1

		// Print Message
		System.out.println("GetWorksheetUniqueId executed successfully.");
	}
}
