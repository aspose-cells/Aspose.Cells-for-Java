package AsposeCellsExamples.Worksheets;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import AsposeCellsExamples.Utils;
public class SetColumnViewWidthInPixels {
	
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CountNumberOfCells.class) + "Worksheets/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");
		
		// Access first worksheet
        Worksheet worksheet = workbook.getWorksheets().get(0);

        // Set the width of the cell in pixels
        worksheet.getCells().setViewColumnWidthPixel(7, 200);

        workbook.save(dataDir + "SetColumnViewWidthInPixels_Out.xls");
        // ExEnd:1

		System.out.println("SetColumnViewWidthInPixels executed successfully.");
	}
}
