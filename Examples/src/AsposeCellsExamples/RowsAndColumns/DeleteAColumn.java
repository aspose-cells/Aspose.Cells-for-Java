package AsposeCellsExamples.RowsAndColumns;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;

public class DeleteAColumn {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(DeleteAColumn.class) + "RowsAndColumns/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "Book1.xlsx");

		// Accessing the first worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Deleting a column from the worksheet at 2nd position
		worksheet.getCells().deleteColumns(1, 1, true);

		// Saving the modified Excel file in default (that is Excel 2000) format
		workbook.save(dataDir + "DeleteAColumn_out.xls");
	}
}
