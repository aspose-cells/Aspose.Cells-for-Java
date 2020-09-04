package AsposeCellsExamples.RowsAndColumns;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;
import AsposeCellsExamples.Worksheets.CountNumberOfCells;

public class SetColumnViewWidthInPixels {
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SetColumnViewWidthInPixels.class) + "RowsAndColumns/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "Book1.xlsx");
		
		// Access first worksheet
        Worksheet worksheet = workbook.getWorksheets().get(0);

        // Set the width of the cell in pixels
        worksheet.getCells().setViewColumnWidthPixel(7, 200);

        workbook.save(dataDir + "SetColumnViewWidthInPixels_Out.xlsx");
        // ExEnd:1

		System.out.println("SetColumnViewWidthInPixels executed successfully.");
	}
}
