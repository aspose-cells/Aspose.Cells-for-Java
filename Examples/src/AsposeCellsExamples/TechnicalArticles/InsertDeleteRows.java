package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;

public class InsertDeleteRows {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(InsertDeleteRows.class) + "TechnicalArticles/";
		// Instantiate a Workbook object.
		Workbook workbook = new Workbook(dataDir + "MyBook.xls");

		// Get the first worksheet in the book.
		Worksheet sheet = workbook.getWorksheets().get(0);

		// Insert 10 rows at row index 2 (insertion starts at 3rd row)
		sheet.getCells().insertRows(2, 10);

		// Delete 5 rows now. (8th row - 12th row)
		sheet.getCells().deleteRows(7, 5, true);

		// Save the Excel file.
		workbook.save(dataDir + "InsertDeleteRows_out.xls");

	}
}
