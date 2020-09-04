package AsposeCellsExamples.RowsAndColumns;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;

public class SummaryRowBelow {
	public static void main(String[] args) throws Exception {
		// ExStart: 1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SummaryRowBelow.class) + "RowsAndColumns/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Accessing the first worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);
		Cells cells = worksheet.getCells();

		// Grouping first six rows (from 0 to 5) and making them hidden by passing true
		cells.groupRows(0, 5, true);

		// Grouping first three columns (from 0 to 2) and making them hidden by passing true
		cells.groupColumns(0, 2, true);

		// Setting SummaryRowBelow property to false
		worksheet.getOutline().setSummaryRowBelow(false);

		// Saving the modified Excel file in default (that is Excel 2003) format
		workbook.save(dataDir + "SummaryRowBelow_out.xls");
		// ExEnd: 1
	}
}
