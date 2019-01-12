package AsposeCellsExamples.Data;

import com.aspose.cells.Range;
import com.aspose.cells.Workbook;
import com.aspose.cells.WorksheetCollection;
import AsposeCellsExamples.Utils;

public class IdentifyCellsinNamedRange {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(IdentifyCellsinNamedRange.class) + "Data/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		WorksheetCollection worksheets = workbook.getWorksheets();

		// Getting the specified named range
		Range range = worksheets.getRangeByName("TestRange");

		// Identify range cells.
		System.out.println("First Row : " + range.getFirstRow());
		System.out.println("First Column : " + range.getFirstColumn());
		System.out.println("Row Count : " + range.getRowCount());
		System.out.println("Column Count : " + range.getColumnCount());

	}
}
