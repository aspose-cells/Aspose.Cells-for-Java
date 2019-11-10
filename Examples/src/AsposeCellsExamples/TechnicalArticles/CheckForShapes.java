package AsposeCellsExamples.TechnicalArticles;

import java.util.Iterator;

import com.aspose.cells.Range;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;

public class CheckForShapes {
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CheckForShapes.class) + "TechnicalArticles/";

		// Create an instance of Workbook and load an existing spreadsheet
		Workbook workbook = new Workbook(dataDir + "SampleCheckCells.xlsx");
		// Loop over all worksheets in the workbook
		for (int i = 0; i < workbook.getWorksheets().getCount(); i++) {
			Worksheet worksheet = workbook.getWorksheets().get(i);
			// Check if worksheet has populated cells
			if (worksheet.getCells().getMaxDataRow() != -1) {
				System.out.println(worksheet.getName() + " is not empty because one or more cells are populated");
			}
			// Check if worksheet has shapes
			else if (worksheet.getShapes().getCount() > 0) {
				System.out.println(worksheet.getName() + " is not empty because there are one or more shapes");
			}
			// Check if worksheet has empty initialized cells
			else {
				Range range = worksheet.getCells().getMaxDisplayRange();
				Iterator rangeIterator = range.iterator();
				if (rangeIterator.hasNext()) {
					System.out.println(worksheet.getName() + " is not empty because one or more cells are initialized");
				} else {
					System.out.println(worksheet.getName() + " is empty");
				}
			}
		}
		// ExEnd:1
	}
}
